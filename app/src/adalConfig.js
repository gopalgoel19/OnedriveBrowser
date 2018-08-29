import { AuthenticationContext, adalFetch, withAdalLogin, adalGetToken } from 'react-adal';

export const adalConfig = {
  clientId: 'b3fc97a4-1c4f-427d-9adb-2c3f3a5aa4d5',
  endpoints: {
    api: 'https://graph.microsoft.com',
  },
  cacheLocation: 'localStorage',
};

export const authContext = new AuthenticationContext(adalConfig);

export const adalApiFetch = (fetch, url, options) =>
  adalFetch(authContext, adalConfig.endpoints.api, fetch, url, options);

export const withAdalLoginApi = withAdalLogin(authContext, adalConfig.endpoints.api);

export const getToken = adalGetToken(authContext, adalConfig.endpoints.api);