import React from 'react';
import { File, FileText  } from 'react-feather';

const FileIcon: React.FC<{ fileType: string }> = ({ fileType }) => {
  let IconComponent;

  // Determine the icon based on the file type
  switch (fileType) {
    case "application/pdf":
      IconComponent = File;
      break;
      case "application/Docx":
      IconComponent = FileText;
      break;
    case "text/plain":
      IconComponent = FileText;
      break;
    default:
      IconComponent = File; // Default icon for unknown types
  }

  return <IconComponent />;
};

export default FileIcon;
