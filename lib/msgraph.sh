# shellcheck shell=bash
# Microsoft Graph API token helpers.
# Standard (delegated) functions call p6df::modules::azure::oauth::token.
# Service-principal (sp::) functions call p6df::modules::azure::auth::sp::token.
# The app registration in Azure Entra must be granted the corresponding
# delegated or application permission for each call to succeed.

# ── Delegated (user OAuth) ───────────────────────────────────────────────────

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::mail::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::mail::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Mail.Read")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::calendar::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::calendar::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Calendars.Read")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::contacts::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::contacts::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Contacts.Read")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::onedrive::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::onedrive::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Files.Read.All")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::sharepoint::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::sharepoint::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Sites.Read.All")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::teams::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::teams::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Team.ReadBasic.All")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::teams::channels::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::teams::channels::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Channel.ReadBasic.All")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::teams::messages::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::teams::messages::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/ChannelMessage.Read.All")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::planner::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::planner::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Tasks.Read")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::onenote::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::onenote::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Notes.Read")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::todo::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::todo::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Tasks.ReadWrite")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::users::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::users::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/User.Read.All")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::groups::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::groups::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Group.Read.All")

  p6_return_str "${token}"
}

######################################################################
#<
#
# Function: str {token} = p6df::modules::m365::directory::get(email)
#
#  Args:
#	email -
#
#  Returns:
#	str - {token}
#
#>
######################################################################
p6df::modules::m365::directory::get() {
  local email="${1:?requires email address}"

  local token
  token=$(p6df::modules::azure::oauth::token "${email}" \
    "https://graph.microsoft.com/Directory.Read.All")

  p6_return_str "${token}"
}

# ── Service-Principal (application permissions) ──────────────────────────────
# Each function mirrors its delegated counterpart but uses client credentials
# (client ID + secret + tenant) via MSAL to acquire an application token.
# The app registration must have the corresponding *application* permission
# (not delegated) granted and admin-consented in Entra ID.
# Scopes are always https://graph.microsoft.com/.default for client credentials.

######################################################################
#<
#
# Function: p6df::modules::m365::sp::mail::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::mail::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::calendar::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::calendar::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::contacts::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::contacts::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::onedrive::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::onedrive::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::sharepoint::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::sharepoint::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::teams::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::teams::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::teams::channels::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::teams::channels::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::teams::messages::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::teams::messages::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::planner::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::planner::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::onenote::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::onenote::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::todo::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::todo::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::users::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::users::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::groups::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::groups::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::directory::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#>
######################################################################
p6df::modules::m365::sp::directory::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}

######################################################################
#<
#
# Function: p6df::modules::m365::sp::all::get(client_id, client_secret, tenant_id)
#
#  Args:
#	client_id -
#	client_secret -
#	tenant_id -
#
#  Notes:
#	Acquires a token with all Graph application permissions granted
#	to the app registration (via .default). Requires admin consent
#	for every permission the app declares in Entra ID.
#
#>
######################################################################
p6df::modules::m365::sp::all::get() {
  local client_id="${1:?requires client ID}"
  local client_secret="${2:?requires client secret}"
  local tenant_id="${3:?requires tenant ID}"

  p6df::modules::azure::auth::sp::token "${client_id}" "${client_secret}" "${tenant_id}" \
    "https://graph.microsoft.com/.default"
}
