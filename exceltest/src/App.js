import React from 'react';
import { Button } from 'office-ui-fabric-react';

const App = () => {
  const writeToCell = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.values = [['Hello, Excel!']];
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };

  return ( 
    <div>
      <h1>Excel Add-In</h1>
      <Button onClick={writeToCell}>Write to Cell</Button>
    </div>
  );
};

export default App;
