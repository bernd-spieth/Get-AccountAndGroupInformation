# Get-AccountAndGroupInformation.ps1
The script collects information about group memberships on servers and clients. This version determines the members in the Remote Desktop Users and Administrators groups. The collected information is then stored in a file share so that it can be analyzed. We use this script to determine which computers and servers have critical permission combinations that need to be eliminated before implementing a 3-tier model in Active Directory.

# Merge-Information.ps1
This script aggregates the individual files stored by the computers in a share and generates a single CSV file from them.

# Deployment
To distribute the script to clients and servers, we use group policies and scheduled tasks. The scheduled tasks make sure that the script is executed several times a day so that the file share has up-to-date information about the local computer groups.
