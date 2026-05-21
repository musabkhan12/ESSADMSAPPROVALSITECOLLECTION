import React from 'react';

// Avatar component that generates initials from a full name
const AvtarComponents: React.FC<{ Name: string }> = ({ Name }) => {
  const avatarStyle: React.CSSProperties = {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    width: '1.5rem', // Adjust size as needed
    height: '1.5rem',
    borderRadius: '50%',
    backgroundColor: '#cfd8dc', // Light gray background
    color: '#546e7a', // Darker gray for initials
    fontSize: '10px', // Adjust font size as needed
    fontWeight: 'bold',
  };

  // Split the name into first and last name
  const nameParts = Name.trim()?.split(' ');
  const firstName = nameParts[0];
  const lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : ''; // Handle single names
  
  // Get initials from first and last name
  const initials = `${firstName[0]}${lastName ? lastName[0] : ''}`.toUpperCase();

  return <div style={avatarStyle}>{initials}</div>;
};

export default AvtarComponents;
