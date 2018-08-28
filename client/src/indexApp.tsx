import * as React from 'react';
import * as ReactDOM from 'react-dom';
import App from './App';
import './index.css';
import { getToken, authContext } from './adalConfig';
import ApolloClient from "apollo-client";
import { ApolloProvider } from "react-apollo";
import { setContext } from 'apollo-link-context';
import { createHttpLink } from 'apollo-link-http';
import { InMemoryCache } from 'apollo-cache-inmemory';
import { runWithAdal, adalGetToken } from 'react-adal';
import { adalConfig } from './adalConfig';

runWithAdal(authContext,()=>{
  adalGetToken(authContext, adalConfig.endpoints.api).then((token)=>{
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
});