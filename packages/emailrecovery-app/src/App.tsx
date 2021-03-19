import React, { useEffect, useState } from 'react';
import { RecoveryComponent } from './features/recovery/RecoveryComponent';

function App() {

  const [ isLoaded, setIsLoaded ] = useState(false);
  useEffect(() => Office.initialize = () => {
    setIsLoaded(true);
  });

  if (!isLoaded) {
    return <div>Loading...</div>;
  }

  return (
    <div className="App">
      <RecoveryComponent />
    </div>
  );
}

export default App;
