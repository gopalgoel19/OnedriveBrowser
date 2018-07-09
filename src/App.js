import React, { Component } from 'react';
// import Calendar from 'react-calendar';
import logo from './logo.svg';
import './App.css';

// 48860666228-v61riqpl3dnp2065j0c3maf72s9kv953.apps.googleusercontent.com
// tiyjsRCxO81ToFK-t71jiU94

const Button = (props) => {

  let handleClick = () => {
    alert(props.func);
  };

  return(
   <span onClick={handleClick}>
     {props.func} 
   </span>
  );
};

const Navigation = (props) => {
  return (
    <div>
        
    </div>
  );
};

const Display = (props) => {
  return (
    <div>
      <Button func='decrease-year'/>
      <Button func='decrease-month'/>
      <Button func='change-date'/>
      <Button func='increase-month'/>
      <Button func='increase-year'/>
    </div>
  );
};

class Calendar extends React.Component{
  state = {
    day: 2,
    month: "July",
    year: 2018
  };

  changeDate = () => {
    this.setState((prevState) => ({
      day: prevState.day + 1
    }));
  };
  
  render(){
    return(
      <div style={{'text-align': 'center'}}>
        <h3>Calender App</h3>
        <p>{this.state.day}, {this.state.month} {this.state.year}</p>
        <Navigation onClick={this.changeDate}/>
        <Display/>
      </div>
    );
  }
}
class App extends Component {
  render() {
    return (
      <div className="App">
        <header className="App-header">
          <img src={logo} className="App-logo" alt="logo" />
          <h1 className="App-title">Welcome to React!</h1>
        </header>
        <Calendar/>
      </div>
    );
  }
}

export default App;
