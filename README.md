# Add Row to Excel Spreadsheet on Sharepoint


This has been incredibly frustrating. 

Ultimately, all I want to do is append a row to this Excel file, on OneDrive: [**Contact Form Test Sheet.xlsx**](https://ideasonpurpose-my.sharepoint.com/:x:/r/personal/iop_ideasonpurpose_com/Documents/Active%20Jobs/IOP/IOP021_%20BNB/Brand%20New%20Brand!%20Cycle%205/a_Mangement/Contact%20Form%20Test%20Sheet.xlsx?d=web3ac602a73a4e21835edfb4cbd490a8&csf=1&web=1&e=iuD8jT)

## Goals:

- [x] Connect to Azure using the [Graph API Console app](https://learn.microsoft.com/en-us/entra/identity-platform/tutorial-v2-nodejs-console) 
- [x] Call other API endpoints (not just [getUsers](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http))
- [x] Get Users' Drives (OneDrive)
- [x] List contents of Users' Drives
- [x] List directory/contents of a Drive/[DriveItem](https://learn.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0)
- [x] List the file or file attributes for the [target Excel sheet](https://ideasonpurpose-my.sharepoint.com/:x:/r/personal/iop_ideasonpurpose_com/Documents/Active%20Jobs/IOP/IOP021_%20BNB/Brand%20New%20Brand!%20Cycle%205/a_Mangement/Contact%20Form%20Test%20Sheet.xlsx?d=web3ac602a73a4e21835edfb4cbd490a8&csf=1&web=1&e=iuD8jT)
- [x] Append a row to an excel sheet

Additional stuff: 
- [x] Replace Axios with native Fetch API
- [x] Actual API error messages 
- [ ] Convert to ES Modules


## References and things to followup on

- The [PnP library](https://pnp.github.io/pnpjs/) might make some of this easier? Or not? 
- Also the [@microsoft/microsoft-graph-client](https://github.com/microsoftgraph/msgraph-sdk-javascript) library
