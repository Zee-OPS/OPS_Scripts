# Deploy a OnePlace Solutions Email Content Type to multiple Site Collections and Document Libraries

A script and sample CSV file to add OnePlace Solutions Email Columns to Content Types in listed Site Collections, create the Content Types where necessary, and add them to listed Document Libraries.

## Getting Started

Download the SitesDocLibs.csv file above and customize it to your requirements.
Note that you need a new line for each uniquely named Site Content Type, and to define which Site Collection it will be created in, and which Document Library it will be added to.
Any Site Content Types listed that already exist in your SharePoint Environment will have the Email Columns added to it (and preserve the existing columns).

### Prerequisites

* [OnePlace Solutions Email Columns](https://github.com/OnePlaceSolutions/EmailColumnsPnP) have been installed to the Site Collections you wish to deploy to.
* Administrator rights to your SharePoint Admin Site and the Site Collections you wish to deploy to.
* [SharePoint PnP CmdLets for SharePoint Online](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps) - Required for executing the modifications against your Site Collections
* [SharePoint Online Management Shell](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps) - Required to Authenticate against your Admin Site and access the listed Site Collections through said authentication.

### Assumptions

* This script assumes that the Content Type(s) to be created will have the Site Content Type 'Document' for it's Parent Content Type. If you are using this script to add the Email Columns to an existing Content Type, this will be updated to inherit from the 'Document' Site Content Type in the process. 

### Restrictions

* Deploying to Sub-Sites/Subwebs using this script is currently unsupported. If you list a Sub-Site or Subweb in the CSV you supply to the script, it will be identified and skipped.

### Installing

1. Start PowerShell on your machine

2. Run the below command to invoke the current version of the script:

```
Invoke-Expression (New-Object Net.WebClient).DownloadString(‘https://raw.githubusercontent.com/ashleygagregory/OPS_Scripts/master/Content%20Type%20Deployment/Multiple%20Sites%20Multiple%20Libraries/DeployECTToSitesDoclibs.ps1’)
```

And repeat

```
until finished
```

End with an example of getting some data out of the system or using it for a little demo

## Running the tests

Explain how to run the automated tests for this system

### Break down into end to end tests

Explain what these tests test and why

```
Give an example
```

### And coding style tests

Explain what these tests test and why

```
Give an example
```

## Deployment

Add additional notes about how to deploy this on a live system

## Built With

* [Dropwizard](http://www.dropwizard.io/1.0.2/docs/) - The web framework used
* [Maven](https://maven.apache.org/) - Dependency Management
* [ROME](https://rometools.github.io/rome/) - Used to generate RSS Feeds


## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Colin Wood for his code example on CSV parsing/iterating

