import React from 'react';
import styles from '../../Form.module.scss';


const AlertComponent:  React.FC = () => {


  return (
    <span  className={`${styles.warning}`}>*mandiatory</span>
  );
};

export default AlertComponent;
