This project has two parts:
- App: It is a SPA(Single Page Application) which was bootstrapped with [Create React App](https://github.com/facebookincubator/create-react-app). For authentication [React ADAL](https://github.com/salvoravida/react-adal) was used.
- QraphQl: It is a GraphQl server which is like a layer above Microsoft Graph Explorer REST API. The App queries this GraphQL server to get all the data needed at frontend.

To run this app locally, we need to start both the App and the GraphQL server. Follow the steps below to run this project:

## App
To run the App locally, run the following command inside app folder:

### `npm install`

Installs the app dependencies specified in the package.json file.<br>

### `npm start`

Runs the app in the development mode.<br>
Open [http://localhost:3000](http://localhost:3000) to view it in the browser.

The page will reload if you make edits.<br>
You will also see any lint errors in the console.

## GraphQL
To run the GraphQL server locally, run the following command inside graphql folder:

### `yarn`

Installs the app dependencies specified in the package.json file.<br>

### `node ./index.js`

Runs the server.
