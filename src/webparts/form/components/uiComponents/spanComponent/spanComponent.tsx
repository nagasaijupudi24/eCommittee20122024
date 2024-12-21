import React from 'react';
import styles from '../../Form.module.scss';


const SpanComponent:  React.FC = () => {


  return (
    <span  className={`${styles.warning}`}>*</span>
  );
};

export default SpanComponent;
