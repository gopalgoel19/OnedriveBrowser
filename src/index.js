import { runWithAdal } from 'react-adal';
import { authContext } from './adalConfig';
import { adalApiFetch } from './adalConfig';

const DO_NOT_LOGIN = false;

runWithAdal(authContext, () => {

  // eslint-disable-next-line
  // console.log(authContext);
  require('./indexApp.js');

},DO_NOT_LOGIN);