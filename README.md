# autoDL
This script provides a clean, automated workaround for Microsoft 365’s inability to retroactively mail-enable existing standard Security Groups (i.e., convert a security group to a mail-enabled security group) — it keeps a Distribution List in sync with group membership, enabling communication without compromising access design.

---

## 🚀 Overview

Microsoft 365 does not support converting existing Security Groups into mail-enabled Security Groups. This limitation presents a challenge for organizations that use Security Groups for access control but also need email distribution capabilities tied to those same groups.

This script bridges that gap by:

- Automatically syncing members from an Azure AD Security Group to an Exchange Online Distribution List
- Creating the DL if it doesn't exist
- Keeping memberships continuously in sync — additions and removals
- Running unattended via Azure Automation and Managed Identity

---

## 🛠️ Requirements

- **Azure Automation Account** with a system-assigned Managed Identity
- **Microsoft.Graph** and **ExchangeOnlineManagement** PowerShell modules installed in the automation environment
- **Automation Variable**: `SecurityGroupCsv` — a CSV string containing one or more Security Group names

---

## 📥 SecurityGroupCsv Variable Format

The script expects a CSV-formatted string stored in an Azure Automation variable named `SecurityGroupCsv`.

### Example value:
```csv
SecurityGroup
Finance-Team
IT-Admins
HR-General
