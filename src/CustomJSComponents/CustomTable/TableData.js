// import React, { useState, useEffect } from 'react';

// const TableData = () => {
//   // Dynamic headers array
//   const headers = ['Entity', 'Level 1', 'Level 2', 'Level 3', 'Level 4', 'Level 5', 'Level 6', 'Level 7', 'Level 8', 'Level 9','Level 10', 'Workflow Type', 'Rework'];

//   // Load initial data from localStorage, or use default
//   const initialData = JSON.parse(localStorage.getItem('tableData')) || [
//     { entity: 'Entity 1', levels: ['', '', '', '', '', '', '', '', '',''], workflowType: '', rework: '' },
//     { entity: 'Entity 2', levels: ['', '', '', '', '', '', '', '', ''], workflowType: '', rework: '' },
//   ];

//   // State to manage table data
//   const [data, setData] = useState(initialData);

//   // Save data to localStorage whenever it changes
//   useEffect(() => {
//     localStorage.setItem('tableData', JSON.stringify(data));
//   }, [data]);

//   // Function to handle changes in the input fields
//   const handleInputChange = (rowIndex, levelIndex, value) => {
//     const updatedData = [...data];
//     updatedData[rowIndex].levels[levelIndex] = value;
//     setData(updatedData);
//   };

//   // Function to handle workflow type and rework input
//   const handleOtherInputChange = (rowIndex, field, value) => {
//     const updatedData = [...data];
//     updatedData[rowIndex][field] = value;
//     setData(updatedData);
//   };

//   // Function to add a new row
//   const addNewRow = () => {
//     const newRow = { entity: `Entity ${data.length + 1}`, levels: [], workflowType: '', rework: '' };
//     setData([...data, newRow]);
//   };

//   // Function to manually save changes (optional)
//   const saveChanges = () => {
//     localStorage.setItem('tableData', JSON.stringify(data));
//     alert('Changes saved!');
//   };

//   return (
//     <div>
//       <h1>Dynamic Workflow Table</h1>
//       <button onClick={addNewRow}>Add New Row</button>
//       <button onClick={saveChanges} style={{ marginLeft: '10px' }}>Save Changes</button>
//       <table border="1" cellPadding="10" style={{ borderCollapse: 'collapse', width: '100%', marginTop: '20px' }}>
//         <thead>
//           <tr>
//             {headers.map((header, index) => (
//               <th key={index}>{header}</th>
//             ))}
//           </tr>
//         </thead>
//         <tbody>
//           {data.map((row, rowIndex) => (
//             <tr key={rowIndex}>
//               <td>{row.entity}</td>
//               {row.levels.map((level, levelIndex) => (
//                 <td key={levelIndex}>
//                   <input
//                     type="text"
//                     value={level}
//                     onChange={(e) => handleInputChange(rowIndex, levelIndex, e.target.value)}
//                     placeholder={`User ${levelIndex + 1}`}
//                   />
//                 </td>
//               ))}
//               <td>
//                 <input
//                   type="text"
//                   value={row.workflowType}
//                   onChange={(e) => handleOtherInputChange(rowIndex, 'workflowType', e.target.value)}
//                   placeholder="Workflow Type"
//                 />
//               </td>
//               <td>
//                 <input
//                   type="text"
//                   value={row.rework}
//                   onChange={(e) => handleOtherInputChange(rowIndex, 'rework', e.target.value)}
//                   placeholder="Rework"
//                 />
//               </td>
//             </tr>
//           ))}
//         </tbody>
//       </table>
//     </div>
//   );
// };

// export default TableData;

import React, { useState, useEffect } from 'react';

const TableData = () => {
  // List of entity options
  const entityOptions = ['Entity 1', 'Entity 2', 'Entity 3', 'Entity 4', 'Entity 5'];

  // Default headers that don't depend on levels
  const staticHeaders = ['Entity', 'Workflow Type', 'Rework'];

  // State for the number of dynamic levels (initially 10)
  const [numLevels, setNumLevels] = useState(10);

  // Load initial data from localStorage, or use default
  const initialData = JSON.parse(localStorage.getItem('tableData')) || [
    { entity: 'Entity 1', levels: Array(numLevels).fill(''), workflowType: '', rework: '' },
    { entity: 'Entity 2', levels: Array(numLevels).fill(''), workflowType: '', rework: '' },
  ];

  // State to manage table data
  const [data, setData] = useState(initialData);

  // Save data to localStorage whenever it changes
  useEffect(() => {
    localStorage.setItem('tableData', JSON.stringify(data));
  }, [data]);

  // Function to handle changes in the input fields for levels
  const handleInputChange = (rowIndex, levelIndex, value) => {
    const updatedData = [...data];
    updatedData[rowIndex].levels[levelIndex] = value;
    setData(updatedData);
  };

  // Function to handle workflow type and rework input
  const handleOtherInputChange = (rowIndex, field, value) => {
    const updatedData = [...data];
    updatedData[rowIndex][field] = value;
    setData(updatedData);
  };

  // Function to handle changes in the entity select dropdown
  const handleEntityChange = (rowIndex, value) => {
    const updatedData = [...data];
    updatedData[rowIndex].entity = value;
    setData(updatedData);
  };

  // Function to add a new row
  const addNewRow = () => {
    const newRow = { entity: '', levels: Array(numLevels).fill(''), workflowType: '', rework: '' };
    setData([...data, newRow]);
  };

  // Function to increase levels
  const increaseLevels = () => {
    setNumLevels(numLevels + 1);
    setData(data.map(row => ({ ...row, levels: [...row.levels, ''] }))); // Add a new empty level for each row
  };

  // Function to decrease levels
  const decreaseLevels = () => {
    if (numLevels > 1) {
      setNumLevels(numLevels - 1);
      setData(data.map(row => ({ ...row, levels: row.levels.slice(0, -1) }))); // Remove the last level from each row
    }
  };

  // Function to manually save changes
  const saveChanges = () => {
    localStorage.setItem('tableData', JSON.stringify(data));
    // alert('Changes saved!');
  };

  return (
    <div>
      <h1>Dynamic Workflow Table</h1>
      <div style={{ marginBottom: '20px' }}>
        <button onClick={addNewRow}>Add New Row</button>
        <button onClick={increaseLevels} style={{ marginLeft: '10px' }}>Add Level</button>
        <button onClick={decreaseLevels} style={{ marginLeft: '10px' }}>Remove Level</button>
        <button onClick={saveChanges} style={{ marginLeft: '10px' }}>Save Changes</button>
      </div>

      <table border="1" cellPadding="10" style={{ borderCollapse: 'collapse', width: '100%' }}>
        <thead>
          <tr>
            <th>Entity</th>
            {[...Array(numLevels)].map((_, index) => (
              <th key={index}>Level {index + 1}</th>
            ))}
            {staticHeaders.slice(1).map((header, index) => (
              <th key={index}>{header}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.map((row, rowIndex) => (
            <tr key={rowIndex}>
              <td>
                <select
                  value={row.entity}
                  onChange={(e) => handleEntityChange(rowIndex, e.target.value)}
                >
                  <option value="">Select Entity</option>
                  {entityOptions.map((entity, entityIndex) => (
                    <option key={entityIndex} value={entity}>
                      {entity}
                    </option>
                  ))}
                </select>
              </td>
              {row.levels.map((level, levelIndex) => (
                <td key={levelIndex}>
                  <input
                    type="text"
                    value={level}
                    onChange={(e) => handleInputChange(rowIndex, levelIndex, e.target.value)}
                    placeholder={`User ${levelIndex + 1}`}
                  />
                </td>
              ))}
              <td>
                <input
                  type="text"
                  value={row.workflowType}
                  onChange={(e) => handleOtherInputChange(rowIndex, 'workflowType', e.target.value)}
                  placeholder="Workflow Type"
                />
              </td>
              <td>
                <input
                  type="text"
                  value={row.rework}
                  onChange={(e) => handleOtherInputChange(rowIndex, 'rework', e.target.value)}
                  placeholder="Rework"
                />
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default TableData;

