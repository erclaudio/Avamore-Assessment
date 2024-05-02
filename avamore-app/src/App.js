import React, { useState, useEffect } from 'react';
import logo from './AVAMORE-LOGO.png';
import './App.css';

function App() {
  const [formValues, setFormValues]= useState({});

  useEffect(() => {  

    fetch('/values')
    .then(res => res.json())
    .then(data => {
      setFormValues(data)
    })
    
  }, []);
function handleSubmit(e) {
  e.preventDefault()
  fetch('/values', {
    method: 'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify(formValues)})
    .then(res => res.json())
    .then(data => {
      setFormValues(data)
    })
  console.log(formValues)
}

  
  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className ="smaller-logo" alt="logo" />
        <p>
          Avamore Loan Tool
        </p>
        <form onSubmit={handleSubmit}>
          <label>
            Facility A Value: <input type ="number" value = {formValues.facility_A} onChange={(e) => setFormValues({...formValues, facility_A: e.target.value})} />
          </label>
          <label>
            Monthly Interest Rate: <input type ="number" step ="" value = {Math.round(formValues.interest_rate*100)/100} onChange={(e) => setFormValues({...formValues, interest_rate: e.target.value})} />
          </label>
          <label>
            Default Period Start : <input type ="date" value = {formValues.default_start} onChange={(e) => setFormValues({...formValues, default_start: e.target.value})} />
          </label>
          <label>
            Default Period End : <input type ="date" value = {formValues.default_end} onChange={(e) => setFormValues({...formValues, default_end: e.target.value})} />
          </label>
          <p><button type ="submit">Submit</button></p>
          
        </form>
        
        {formValues.interest_due !== null && <p>Interest Due: £{formValues.interest_due}</p>}

      </header>
      
    </div>
  );
}

export default App;