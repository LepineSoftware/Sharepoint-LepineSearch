import * as React from 'react';

export interface IHighlightTextProps {
  text: string;
  query: string;
}

export const HighlightText: React.FunctionComponent<IHighlightTextProps> = (props) => {
  const { text, query } = props;

  // If no query, just return the text
  if (!query || !text) {
    return <span>{text}</span>;
  }

  // Split text by the query (case-insensitive) capturing the delimiter
  const regex = new RegExp(`(${query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
  const parts = text.split(regex);

  return (
    <span>
      {parts.map((part, index) => 
        // Check if this part matches the query (case-insensitive check)
        part.toLowerCase() === query.toLowerCase() ? (
          <span key={index} style={{ fontWeight: 'bold', color: '#0078d4' }}>{part}</span>
        ) : (
          <span key={index}>{part}</span>
        )
      )}
    </span>
  );
};