# P6's POSIX.2: p6df-m365

## Table of Contents

- [Badges](#badges)
- [Summary](#summary)
- [Contributing](#contributing)
- [Code of Conduct](#code-of-conduct)
- [Usage](#usage)
  - [Functions](#functions)
- [Hierarchy](#hierarchy)
- [Author](#author)

## Badges

[![License](https://img.shields.io/badge/License-Apache%202.0-yellowgreen.svg)](https://opensource.org/licenses/Apache-2.0)

## Summary

TODO: Add a short summary of this module.

## Contributing

- [How to Contribute](<https://github.com/p6m7g8-dotfiles/.github/blob/main/CONTRIBUTING.md>)

## Code of Conduct

- [Code of Conduct](<https://github.com/p6m7g8-dotfiles/.github/blob/main/CODE_OF_CONDUCT.md>)

## Usage

### Functions

#### p6df-m365

##### p6df-m365/init.zsh

- `p6df::modules::m365::deps()`
- `p6df::modules::m365::external::brews()`
- `p6df::modules::m365::langs()`
- `p6df::modules::m365::mcp()`
- `words m365 = p6df::modules::m365::profile::mod()`

#### p6df-m365/lib

##### p6df-m365/lib/msgraph.sh

- `p6df::modules::m365::licenses::list()`
- `p6df::modules::m365::sp::all::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::calendar::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::contacts::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::directory::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::groups::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::mail::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::onedrive::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::onenote::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::planner::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::sharepoint::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::teams::channels::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::teams::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::teams::messages::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::todo::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::sp::users::get(client_id, client_secret, tenant_id)`
  - Args:
    - client_id
    - client_secret
    - tenant_id
- `p6df::modules::m365::users::list()`
- `str {token} = p6df::modules::m365::calendar::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::contacts::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::directory::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::groups::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::mail::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::onedrive::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::onenote::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::planner::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::sharepoint::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::teams::channels::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::teams::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::teams::messages::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::todo::get(email)`
  - Args:
    - email
- `str {token} = p6df::modules::m365::users::get(email)`
  - Args:
    - email

## Hierarchy

```text
.
├── init.zsh
├── lib
│   ├── licenses.py
│   ├── msgraph.sh
│   └── users.py
└── README.md

2 directories, 5 files
```

## Author

Philip M. Gollucci <pgollucci@p6m7g8.com>
