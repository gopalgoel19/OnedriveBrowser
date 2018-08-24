import * as React from 'react';
import * as ReactDOM from 'react-dom';
import App from './App';
import './index.css';
import { getToken } from './adalConfig';
import registerServiceWorker from './registerServiceWorker';
import gql from "graphql-tag";
import ApolloClient from "apollo-client";
import { ApolloProvider } from "react-apollo";
import { setContext } from 'apollo-link-context';
import { createHttpLink } from 'apollo-link-http';
import { InMemoryCache } from 'apollo-cache-inmemory';

getToken.then((token)=>{
  const authLink = setContext((_, { headers }) => {
    return {
      headers: {
        ...headers,
        authorization: token ? `Bearer ${token}` : ''
      }
    }
  });

  const httpLink = createHttpLink({
    uri: 'http://localhost:4000'
  });

  const client = new ApolloClient({
    // link: httpLink,
    link: authLink.concat(httpLink),
    cache: new InMemoryCache()
  });
  
  ReactDOM.render(
    <ApolloProvider client={client}>
      <App />
    </ApolloProvider>,
    document.getElementById('root') as HTMLElement
  );
});

// console.log(getMyToken());

// const client = new ApolloClient({
//   uri: "http://localhost:4000"
// });

// ReactDOM.render(
//   <ApolloProvider client={client}>
//     <App />
//   </ApolloProvider>,
//   document.getElementById('root') as HTMLElement
// );
registerServiceWorker();
