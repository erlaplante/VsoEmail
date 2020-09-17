PowerShell template to use the VSO (Azure DevOps/TFS) REST API for queries and custom formatting to an Outlook message.

Key Benefits:

- Allows for conditional modification of VSO query, inherent to running as a script rather than using the static web interface.

   - Also removes need for queries in 'Shared Queries' path due to query being built and executed by the script.

- Easily add Work Item URLs to custom query output.

- Extendable template can be used for multiple queries.

- EnhancedHTML2 module provides wealth of html, css, and javascript formatting options for specific use cases.

- PowerShell Credential Manager provides easy processing of PAT for API authentication.

For additional script information:

Get-Help .\Get-VsoEmail.ps1 -Full

### Example

![alt text](https://github.com/erlaplante/VsoEmail/blob/master/exampleOutput.png?raw=true)
