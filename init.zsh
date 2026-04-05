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
# Function: p6df::modules::m365::external::brews()
#
#>
######################################################################
p6df::modules::m365::external::brews() {

  p6df::core::homebrew::cli::brew::install --cask microsoft-teams

  p6_return_void
}

######################################################################
#<
#
# Function: p6df::modules::m365::langs()
#
#>
######################################################################
p6df::modules::m365::langs() {

  uv pip install msgraph-sdk msal httpx

  p6_return_void
}

######################################################################
#<
#
# Function: p6df::modules::m365::mcp()
#
#>
######################################################################
p6df::modules::m365::mcp() {

  claude mcp add --transport http "claude.ai Microsoft 365" "https://mcp.microsoft365.com/mcp"

  p6_return_void
}

######################################################################
#<
#
# Function: words m365 = p6df::modules::m365::profile::mod()
#
#  Returns:
#	words - m365
#
#  Environment:	 AZURE_CLIENT_ID AZURE_CLIENT_SECRET AZURE_TENANT_ID
#>
######################################################################
p6df::modules::m365::profile::mod() {

  p6_return_words 'm365' '$AZURE_CLIENT_ID' '$AZURE_CLIENT_SECRET' '$AZURE_TENANT_ID'
}
