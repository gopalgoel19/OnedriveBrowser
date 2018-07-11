import { runWithAdal } from 'react-adal';
import { authContext } from './adalConfig';

const DO_NOT_LOGIN = false;

document.getElementById('hello').textContent = "";

runWithAdal(authContext, () => {

  // eslint-disable-next-line
  // console.log(authContext);
  require('./indexApp.tsx');

},DO_NOT_LOGIN);