import * as React from 'react';
import styles from '../../Form.module.scss';



const Header: React.FC = () => {
  
    

  return (
    <div className={`${styles.headerContainer} ${styles.commonProperties}`}>
      <h1 className={`${styles.headers}`}>Bank Image</h1>
      <h1 style={{"fontSize":"22px"}}>IB Smart Office - eCommittee</h1>
      <div>
        <h5 >Name</h5>
        <img src=""/>
        <button type="button" className={`${styles.commonBtn2}  ${styles.commonBtn}`}>Logout</button>
      </div>
      
      </div>
   
  );
};

export default Header;
