#!/bin/bash

# gets information for all webhooks for in an organization

# need: `gh auth login -h github.com` and auth with a PAT!
# since the Oauth token can only receive results for hooks it created for this API call

# note: tsv is the default format
# tsv is a subset of fields, json is all fields

if [ $# -lt 1 ]
  then
    echo "usage: $0 <enterprise slug> <hostname> <format: tsv|json> > output.tsv/json"
    exit 1
fi

enterpriseslug=$1
hostname=$2
format=$3
export PAGER=""

# set hostname to github.com by default
if [ -z "$hostname" ]
then
  hostname="github.com"
fi

auth_status=$(gh auth token -h $hostname 2>&1)

if [[ $auth_status == gho_* ]]
then
  echo "Token starts with gho_ - use "gh auth login" and authenticate with a PAT with read:org and admin:org_hook scope"
  exit 1
fi
if [ -z "$format" ]
then
  format="tsv"
fi

organizations=$(gh api graphql --hostname $hostname --paginate -f enterpriseName="$enterpriseslug" -f query='
query getEnterpriseOrganizations($enterpriseName: String! $endCursor: String) {
  enterprise(slug: $enterpriseName) {
    organizations(first: 100, after: $endCursor) {
      nodes {
        id
        login
      }
      pageInfo {
        endCursor
        hasNextPage
      }
    }
  }
}' --jq '.data.enterprise.organizations.nodes[].login')

# check to see if organizations is null - null error message is confusing otherwise
if [ -z "$organizations" ] || [ $? -ne 0 ]
then
  # Define color codes
  RED='\033[0;31m'
  NC='\033[0m' # No Color

  # Print colored messages
  echo -e "${RED}No organizations found for enterprise: $enterpriseslug${NC}"
  echo -e "${RED}Check that you have the proper scopes for enterprise, e.g.: 'gh auth refresh -h github.com -s read:org -s read:enterprise'${NC}"
  exit 1
fi

if [ "$format" == "tsv" ]; then
  echo -e "Organization\tActive\tURL\tCreated At\tUpdated At\tEvents"
fi

for org in $organizations
do
  if [ "$format" == "tsv" ]; then
    gh api "orgs/$org/hooks" --hostname $hostname --paginate --jq ".[] | [\"$org\",.active,.config.url, .created_at, .updated_at, (.events | join(\",\"))] | @tsv"
  else
    gh api "orgs/$org/hooks" --hostname $hostname --paginate --jq ".[] | {organization: \"$org\", active: .active, url: .config.url, created_at: .created_at, updated_at: .updated_at, events: .events}"
  fi
done
