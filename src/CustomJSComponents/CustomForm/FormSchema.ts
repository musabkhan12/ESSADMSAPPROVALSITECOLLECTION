// FormSchema.ts
export interface FormField {
    type: string;
    name: string;
    label: string;
    placeholder?: string;
    options?: []; // For select, radio buttons, etc.
  }
  
  export const formSchema: FormField[] = [
    {
      type: "text",
      name: "firstName",
      label: "First Name",
      placeholder: "Enter your first name",
    },
    {
      type: "email",
      name: "email",
      label: "Email",
      placeholder: "Enter your email",
    },
    {
      type: "select",
      name: "gender",
      label: "Gender",
      options:[],
    },
    {
      type: 'radio',
      name: 'gender',
      label: 'Gender',
      options: [
      ],
    },
  ];
  