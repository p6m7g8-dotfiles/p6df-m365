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
# Function: str str = p6df::modules::m365::prompt::mod()
#
#  Returns:
#	str - str
#
#  Environment:	 AZURE_CLIENT_ID AZURE_TENANT_ID P6_DFZ_PROFILE_M365
#>
######################################################################
p6df::modules::m365::prompt::mod() {

  local str
  if p6_string_blank_NOT "$P6_DFZ_PROFILE_M365"; then
    str="m365:\t\t  $P6_DFZ_PROFILE_M365"
    if p6_string_blank_NOT "$AZURE_CLIENT_ID"; then
      str=$(p6_string_append "$str" "client" "/")
    fi
    if p6_string_blank_NOT "$AZURE_TENANT_ID"; then
      str=$(p6_string_append "$str" "tenant" "/")
    fi
  fi

  p6_return_str "$str"
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
# Function: p6df::modules::m365::profile::on(profile, client_id, tenant_id, client_secret)
#
#  Args:
#	profile -
#	client_id -
#	tenant_id -
#	client_secret -
#
#  Environment:	 AZURE_CLIENT_ID AZURE_CLIENT_SECRET AZURE_TENANT_ID P6_DFZ_PROFILE_M365
#>
######################################################################
p6df::modules::m365::profile::on() {
  local profile="$1"
  local client_id="$2"
  local tenant_id="$3"
  local client_secret="${4:-}"

  p6_env_export "P6_DFZ_PROFILE_M365" "$profile"
  p6_env_export "AZURE_CLIENT_ID" "$client_id"
  p6_env_export "AZURE_TENANT_ID" "$tenant_id"
  if p6_string_blank_NOT "$client_secret"; then
    p6_env_export "AZURE_CLIENT_SECRET" "$client_secret"
  fi

  p6_return_void
}

######################################################################
#<
#
# Function: p6df::modules::m365::profile::off()
#
#  Environment:	 AZURE_CLIENT_ID AZURE_TENANT_ID P6_DFZ_PROFILE_M365
#>
######################################################################
p6df::modules::m365::profile::off() {

  p6_env_export_un P6_DFZ_PROFILE_M365
  p6_env_export_un AZURE_CLIENT_ID
  p6_env_export_un AZURE_CLIENT_SECRET
  p6_env_export_un AZURE_TENANT_ID

  p6_return_void
}
