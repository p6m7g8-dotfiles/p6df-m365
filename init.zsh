# shellcheck shell=bash
######################################################################
#<
#
# Function: p6df::modules::m365::deps()
#
#>
######################################################################
p6df::modules::m365::deps() {
  ModuleDeps=(
    p6m7g8-dotfiles/p6df-azure
    microsoftgraph/msgraph-sdk-python
  )
}

######################################################################
#<
#
# Function: p6df::modules::m365::langs()
#
#>
######################################################################
p6df::modules::m365::langs() {

  pip install msgraph-sdk msal httpx

  p6_return_void
}

######################################################################
#<
#
# Function: words m365 $AZURE_CLIENT_ID = p6df::modules::m365::profile::mod()
#
#  Returns:
#	words - m365 $AZURE_CLIENT_ID
#
#  Environment:	 AZURE_CLIENT_ID
#>
######################################################################
p6df::modules::m365::profile::mod() {

  p6_return_words 'm365' "$AZURE_CLIENT_ID"
}
