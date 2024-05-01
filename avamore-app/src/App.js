import React, { useState, useEffect } from 'react';
import logo from './AVAMORE-LOGO.png';
import './App.css';

function App() {
  const [currentTime, setCurrentTime] = useState(0);
  const [facilityValue, setFacilityValue] = useState(null);
  const [interestDue, setInterestDue] = useState(null);
  const [interestRate, setInterestRate] = useState(null);
  const [defaultStart, setDefaultStart] = useState(null)
  const [defaultEnd, setDefaultEnd] = useState(null)

  useEffect(() => {  

    fetch('/values')
    .then(res => res.json())
    .then(data => console.log(data))
    
    fetch('/interest-rate')
    .then(res => res.json())
    .then(data => setInterestRate(data.interest_rate))
    .catch(error => console.error('Error fetching interest rate:', error))

    fetch('/time')
      .then(res => res.json())
      .then(data => setCurrentTime(data.time));

    fetch('/facility')
      .then(res => res.json())
      .then(data => setFacilityValue(data.facility_value))
      .catch(error => console.error('Error fetching facility value:', error));

    fetch('/interest-due')
      .then(res => res.json())
      .then(data => setInterestDue(data.interest_due))
      .catch(error => console.error('Error fetching interest due value:', error));
    
    fetch('/default-period')
      .then(res => res.json())
      .then(data => setDefaultStart(data.default_start))
      .catch(error => console.error('Error fetching default date:', error));
    
    fetch('/default-period')
    .then(res => res.json())
    .then(data => setDefaultEnd(data.default_end))
    .catch(error => console.error('Error fetching default date:', error));
  }, []);

  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className ="smaller-logo" alt="logo" />
        <p>
          Avamore Loan Tool
        </p>
        {facilityValue !== null && <p>Facility A Value: £{facilityValue}</p>}
        {interestRate !== null && <p>Monthly Interest Rate: {interestRate}% </p>}
        {defaultStart !== null && <p>Default Start Date: {defaultStart}</p>}
        {defaultEnd !== null && <p>Default End Date: {defaultEnd}</p>}
        {interestDue !== null && <p>Interest Due: £{interestDue}</p>}

      </header>
      
    </div>
  );
}

export default App;