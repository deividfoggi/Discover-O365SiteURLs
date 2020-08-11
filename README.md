# OneDrive-URL-Finder

This script is intended to make it easier for eDiscovery Data investigars to find User's OneDrive URL to use it as input on content search and eDiscovery cases. Currently, Office 365 content search feature doesn't include an OneDrive URL search by user name, displayname, UPN, etc., making it hard for non-IT investigators to getting URL to scope/narrow down their searches.

To use this tool, you have to register an app in Azure AD of your tenant with Application Permission for Microsoft Graph with permission Files.Read.All. Input the application ID, Tenant ID and client secret on lines 15, 16 and 17 in the respective variables.

The usage is very simple, eDiscovery investigators may run this script and type on the text box below "Insert user's UPN/email" the UPN/e-mail of the user they want to find the OneDrive URL. Then you just have to copy the result from "OneDrive URL" field and paste into content search filter to include this user's OneDrive and narrow down your search.
