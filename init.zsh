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
# Function: p6df::modules::m365::init(_module, dir)
#
#  Args:
#	_module -
#	dir -
#
#>
######################################################################
p6df::modules::m365::init() {
  local _module="$1"
  local dir="$2"

  p6_bootstrap "$dir"

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

  pip install msgraph-sdk msal

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

  p6_js_npm_global_install "@leonardocrdso/office365-mcp-server"

  p6_return_void
}
