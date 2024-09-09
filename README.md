# Model Driven App Cloner

This project provides a tool for cloning Model-Driven Apps in Microsoft Dataverse (Power Apps). It allows you to create a copy of an existing app, including its components and site map, and place it in a new solution.

## Features

- Clone an existing Model-Driven App
- Create a new solution for the cloned app
- Copy and recreate the site map
- Add all components from the source app to the cloned app
- Publish the newly created app

## Prerequisites

- .NET 5.0 or later
- Access to a Microsoft Dataverse environment
- Appropriate permissions to create and modify apps and solutions

## Setup

1. Clone this repository
2. Open the solution in Visual Studio
3. Update the `ConnectToDataverse` method with your Dataverse environment details:
   - Client ID
   - Client Secret
   - Environment URL

## Usage

1. Run the application
2. Enter the name of the source app you want to clone
3. Provide a name for the new cloned app
4. Enter a name for the new solution that will contain the cloned app
5. The application will create the new app, clone all components, and provide you with the URL of the new app

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on our code of conduct, and the process for submitting pull requests to us.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

- Microsoft Dataverse SDK
- Power Apps community
