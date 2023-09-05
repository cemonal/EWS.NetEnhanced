# EWS.NetEnhanced

EWS.NetEnhanced is an enhanced version of the Exchange Web Services library for .NET Standard 2.1. It aims to provide an improved experience and additional features on top of the original EWS Managed API by Microsoft.

This library is based on the [EWS Managed API](https://github.com/OfficeDev/ews-managed-api), with enhancements and updates to make it compatible with .NET Standard 2.1 and modern development practices.

## Getting Started

### Cloning the Repository

To clone this repository, use the following command:

```bash
git clone https://github.com/cemonal/EWS.NetEnhanced.git
```

### Original Code Repository

The original EWS Managed API repository by Microsoft can be found at:

[https://github.com/OfficeDev/ews-managed-api](https://github.com/OfficeDev/ews-managed-api)

### Microsoft Documentation

For detailed information about the Exchange Web Services and the EWS Managed API, refer to the official Microsoft documentation:

[Exchange Web Services Managed API Reference](https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/ews-managed-api-reference-for-exchange)

## Features

- Improved compatibility with .NET Standard 2.1
- Enhanced functionality and bug fixes compared to the original library
- Updated dependencies and modern development practices

## Usage Example

Here's an example of how you can use the `EwsCallerService` class to fetch emails with attachments:

```csharp
using EWS.NetEnhanced;

// Create an EWS client
var service = new EwsCallerService(config.AppId, config.ClientSecret, config.TenantId, config.Scopes, config.EWSUrl, config.Email);

// Fetch emails with attachments
var mailsToReceive = service.GetEmails(request.Date, config.PageSize, true);
```

Please replace the values in the `config` object with your actual configuration settings.

## Contributions

Contributions to this project are welcome! If you find any bugs, have feature requests, or want to contribute improvements, please feel free to open an issue or submit a pull request.

---

**Note:** This project is not affiliated with or endorsed by Microsoft. It is an independent effort to enhance the EWS Managed API for modern development environments.
