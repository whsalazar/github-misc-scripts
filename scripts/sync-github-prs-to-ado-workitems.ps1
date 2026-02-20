<#
.SYNOPSIS
  Backfills Azure Boards work item links for migrated GitHub PRs by:
  - finding AB#<id> references in GitHub PRs
  - removing old Azure DevOps PR links from the work item
  - adding a new hyperlink relation to the GitHub PR URL

.REQUIREMENTS
  - PowerShell 7+
  - A GitHub token with repo read access (GH_PAT)
  - An Azure DevOps PAT with Work Items (read/write) access (ADO_PAT)

.EXAMPLE
  pwsh ./sync-github-prs-to-ado-workitems.ps1 `
    -GitHubOwner "myorg" -GitHubRepo "myrepo" `
    -AzureDevOpsOrg "myadoorg" -AzureDevOpsProject "MyProject" `
    -FromPr 1 -ToPr 99999 `
    -DryRun

  Remove -DryRun to apply changes.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $GitHubOwner,
  [Parameter(Mandatory)] [string] $GitHubRepo,

  [Parameter(Mandatory)] [string] $AzureDevOpsOrg,
  [Parameter(Mandatory)] [string] $AzureDevOpsProject,

  # GitHub API token (classic PAT or fine-grained token). Can also be provided via env:GH_PAT
  [string] $GitHubToken = $env:GH_PAT,

  # Azure DevOps PAT with Work Items read/write. Can also be provided via env:ADO_PAT
  [string] $AzureDevOpsPat = $env:ADO_PAT,

  # PR range control (optional)
  [int] $FromPr = 1,
  [int] $ToPr = 2147483647,

  # Where to search for AB# references
  [switch] $ScanTitleAndBody = $true,
  [switch] $ScanComments = $true,

  # Remove relations that look like old ADO PR artifacts (vstfs PR URIs)
  [switch] $RemoveAdoGitPullRequestRelations = $true,

  # Additionally remove hyperlinks that match this regex (e.g., old ADO PR web URLs)
  # Example: '^https://dev\.azure\.com/myadoorg/.+/_git/.+/pullrequest/\d+'
  [string] $RemoveHyperlinksMatchingRegex = "",

  # Add GitHub PR link as Hyperlink relation
  [switch] $AddGitHubPrHyperlink = $true,

  # Dry run mode: prints intended changes only
  [switch] $DryRun
)

if (-not $GitHubToken) { throw "GitHubToken missing. Provide -GitHubToken or set GH_PAT." }
if (-not $AzureDevOpsPat) { throw "AzureDevOpsPat missing. Provide -AzureDevOpsPat or set ADO_PAT." }

# ---------- helpers ----------
function New-BasicAuthHeader([string] $pat) {
  $bytes = [Text.Encoding]::ASCII.GetBytes(":$pat")
  $b64 = [Convert]::ToBase64String($bytes)
  return @{ Authorization = "Basic $b64" }
}

function Invoke-GitHub([string] $method, [string] $url, $body = $null) {
  $headers = @{
    Authorization = "Bearer $GitHubToken"
    Accept        = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
    "User-Agent"  = "gh-ado-linker"
  }
  if ($body) {
    return Invoke-RestMethod -Method $method -Uri $url -Headers $headers -Body ($body | ConvertTo-Json -Depth 20) -ContentType "application/json"
  }
  return Invoke-RestMethod -Method $method -Uri $url -Headers $headers
}

function Invoke-Ado([string] $method, [string] $url, $body = $null, [string] $contentType = "application/json") {
  $headers = New-BasicAuthHeader $AzureDevOpsPat
  if ($body) {
    return Invoke-RestMethod -Method $method -Uri $url -Headers $headers -Body ($body | ConvertTo-Json -Depth 50) -ContentType $contentType
  }
  return Invoke-RestMethod -Method $method -Uri $url -Headers $headers
}

function Get-AbWorkItemIdsFromText([string] $text) {
  if (-not $text) { return @() }
  # Matches AB#123, ab#123, AB # 123
  $regex = [regex] '(?i)\bAB\s*#\s*(\d+)\b'
  $ids = New-Object System.Collections.Generic.HashSet[int]
  foreach ($m in $regex.Matches($text)) {
    [void] $ids.Add([int]$m.Groups[1].Value)
  }
  return $ids.ToArray() | Sort-Object
}

function Get-AdoWorkItem([int] $id) {
  $url = "https://dev.azure.com/$AzureDevOpsOrg/$AzureDevOpsProject/_apis/wit/workitems/$id?`$expand=relations&api-version=7.1"
  return Invoke-Ado GET $url
}

function Patch-AdoWorkItemRelations([int] $id, [object[]] $ops) {
  if (-not $ops -or $ops.Count -eq 0) { return $null }

  if ($DryRun) {
    Write-Host "DRYRUN: Would PATCH work item $id with ops:" -ForegroundColor Yellow
    $ops | ConvertTo-Json -Depth 50 | Write-Host
    return $null
  }

  $url = "https://dev.azure.com/$AzureDevOpsOrg/$AzureDevOpsProject/_apis/wit/workitems/$id?api-version=7.1"
  # JSON Patch content type
  return Invoke-Ado PATCH $url $ops "application/json-patch+json"
}

function Build-RemoveRelationOps($relations, [switch] $removeVstfsPr, [string] $removeHyperlinkRegex) {
  $ops = @()

  if (-not $relations) { return $ops }

  $rx = $null
  if ($removeHyperlinkRegex) { $rx = [regex] $removeHyperlinkRegex }

  for ($i = 0; $i -lt $relations.Count; $i++) {
    $rel = $relations[$i]
    $relUrl = [string]$rel.url
    $relType = [string]$rel.rel

    $remove = $false

    if ($removeVstfsPr) {
      # Typical ADO Git PR artifact URI: vstfs:///Git/PullRequestId/{projectId}%2F{repoId}%2F{prId}
      if ($relUrl -match '^vstfs:\/\/\/Git\/PullRequestId\/') { $remove = $true }
    }

    if (-not $remove -and $rx -and $relType -eq 'Hyperlink') {
      if ($rx.IsMatch($relUrl)) { $remove = $true }
    }

    if ($remove) {
      $ops += @{
        op   = "remove"
        path = "/relations/$i"
      }
    }
  }

  # IMPORTANT: When removing multiple items by index in JSON Patch, remove from highest index to lowest
  # so earlier removals don't shift indices.
  if ($ops.Count -gt 1) {
    $ops = $ops | Sort-Object { [int]($_.path -replace '^/relations/', '') } -Descending
  }

  return $ops
}

function HasExistingGitHubPrHyperlink($relations, [string] $prUrl) {
  if (-not $relations) { return $false }
  foreach ($rel in $relations) {
    if ($rel.rel -eq "Hyperlink" -and [string]$rel.url -eq $prUrl) { return $true }
  }
  return $false
}

function Build-AddGitHubPrOp([string] $prUrl, [string] $comment) {
  return @(
    @{
      op    = "add"
      path  = "/relations/-"
      value = @{
        rel  = "Hyperlink"
        url  = $prUrl
        attributes = @{
          comment = $comment
        }
      }
    }
  )
}

# ---------- enumerate PRs ----------
# GitHub pulls API supports pagination. We'll page by updated desc to make it manageable.
$perPage = 100
$page = 1
$done = $false

while (-not $done) {
  $listUrl = "https://api.github.com/repos/$GitHubOwner/$GitHubRepo/pulls?state=all&sort=updated&direction=desc&per_page=$perPage&page=$page"
  $prs = Invoke-GitHub GET $listUrl

  if (-not $prs -or $prs.Count -eq 0) { break }

  foreach ($pr in $prs) {
    $prNumber = [int]$pr.number
    if ($prNumber -lt $FromPr -or $prNumber -gt $ToPr) { continue }

    $prUrl = [string]$pr.html_url
    $title = [string]$pr.title
    $body  = [string]$pr.body

    $workItemIds = @()

    if ($ScanTitleAndBody) {
      $workItemIds += Get-AbWorkItemIdsFromText $title
      $workItemIds += Get-AbWorkItemIdsFromText $body
    }

    if ($ScanComments) {
      $commentsUrl = "https://api.github.com/repos/$GitHubOwner/$GitHubRepo/issues/$prNumber/comments?per_page=100"
      $comments = Invoke-GitHub GET $commentsUrl
      foreach ($c in $comments) {
        $workItemIds += Get-AbWorkItemIdsFromText ([string]$c.body)
      }
    }

    $workItemIds = $workItemIds | Sort-Object -Unique

    if ($workItemIds.Count -eq 0) { continue }

    Write-Host "PR #$prNumber -> Work items: $($workItemIds -join ', ')  ($prUrl)" -ForegroundColor Cyan

    foreach ($wid in $workItemIds) {
      try {
        $wi = Get-AdoWorkItem -id $wid
      } catch {
        Write-Warning "Failed to fetch work item $wid. $_"
        continue
      }

      $relations = $wi.relations

      $ops = @()

      # 1) Remove old ADO PR relations / hyperlinks
      $ops += Build-RemoveRelationOps -relations $relations `
        -removeVstfsPr:([bool]$RemoveAdoGitPullRequestRelations) `
        -removeHyperlinkRegex $RemoveHyperlinksMatchingRegex

      # 2) Add new GitHub PR hyperlink if not already present
      if ($AddGitHubPrHyperlink) {
        if (-not (HasExistingGitHubPrHyperlink -relations $relations -prUrl $prUrl)) {
          $ops += Build-AddGitHubPrOp -prUrl $prUrl -comment "Migrated PR link (GitHub PR #$prNumber) from $GitHubOwner/$GitHubRepo"
        } else {
          Write-Host "Work item $wid already has GitHub PR hyperlink." -ForegroundColor DarkGray
        }
      }

      if ($ops.Count -eq 0) {
        Write-Host "No changes needed for work item $wid." -ForegroundColor DarkGray
        continue
      }

      try {
        $result = Patch-AdoWorkItemRelations -id $wid -ops $ops
        if (-not $DryRun) {
          Write-Host "Updated work item $wid successfully." -ForegroundColor Green
        }
      } catch {
        Write-Warning "Failed to update work item $wid for PR #$prNumber. $_"
      }
    }
  }

  $page++
  # safety stop to avoid infinite loops in case of unexpected API behavior
  if ($page -gt 1000) { $done = $true }
}
