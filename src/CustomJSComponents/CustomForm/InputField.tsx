// import * as React from 'react';
// import { FormField } from './FormSchema';

// interface InputFieldProps {
//   field: FormField;
//   value: string;
//   onChange: (name: string, value: string) => void;
// }

// export const InputField: React.FC<InputFieldProps> = ({ field, value, onChange }) => {
//   const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
//     onChange(field.name, e.target.value);
//   };

//   switch (field.type) {
//     case 'text':
//     case 'email':
//       return (
//         <div>
//           <label>{field.label}</label>
//           <input
//             type={field.type}
//             name={field.name}
//             placeholder={field.placeholder}
//             value={value}
//             onChange={handleChange}
//           />
//         </div>
//       );

//     case 'textarea':
//       return (
//         <div>
//           <label>{field.label}</label>
//           <textarea
//             name={field.name}
//             placeholder={field.placeholder}
//             value={value}
//             onChange={handleChange}
//           />
//         </div>
//       );
//       case 'radio':
//         return (
//           <div className="mb-3">
//             <label>{field.label}</label>
//             {field.options?.map(option => (
//               <div className="form-check">
//                 <input
//                   type="radio"
//                   name={field.name}
//                   value={option.id}
//                   checked={value === option.id}
//                   onChange={handleChange}
//                   className="form-check-input"
//                 />
//                 <label className="form-check-label">{option.name}</label>
//               </div>
//             ))}
//           </div>
//         );
//       // case 'radio': 
//       // return(
//       //   field.options?.map(option => (
//       //     <div key={option.id}>
//       //       <input
//       //         type="radio"
//       //         name={field.name}
//       //         value={option.id}
//       //         onChange={(e) => onChange(field.name, e.target.value)}
//       //       />
//       //       <label>{option.name}</label>
//       //     </div>
//       //   ))
//       // )
//     case 'select':
//       return (
//         <div>
//           <label>{field.label}</label>
//           <select name={field.name} value={value} onChange={handleChange}>
//             {field.options?.map((option:any) => (
//               <option key={option.id} value={option.id}>
//                 {option.name}
//               </option>
//             ))}
//           </select>
//         </div>
//       );
     
//     default:
//       return null;
//   }
// };
