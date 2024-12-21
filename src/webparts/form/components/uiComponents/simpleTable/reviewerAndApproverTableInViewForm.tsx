/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/ban-ts-comment */
import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, IDetailsListStyles, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { format } from 'date-fns';
import { Icon } from '@fluentui/react';

const detailsListStyles: Partial<IDetailsListStyles> = {
    root: {
      paddingTop: '0px', 
    },
  };

const ApproverAndReviewerTableInViewForm = (props: any) => {
    const { type } = props;
    const gridData = props.data;

  
    const columns: IColumn[] = [
        
        { key: 'text', name: type, fieldName: 'text', minWidth: 60, maxWidth: 120, isResizable: true },
        { key: 'srNo', name: 'SR No', fieldName: 'srNo', minWidth: 60, maxWidth: 120, isResizable: true },
        { key: 'optionalText', name: 'Designation', fieldName: 'optionalText', minWidth: 80, maxWidth: 150, isResizable: true },
        {
          key: 'status',
          name: 'Status',
          fieldName: 'status',
          minWidth: 100,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: any) => {
           
        
            let iconName = '';
          
            switch (item.statusNumber) {
              case "2000": 
              case "3000": 
                iconName = 'AwayStatus';
                break;
             
              case '4000':
                iconName = 'Forward';
                break;
              case '6000':
                iconName = 'Reply';
                break;
              case '8000':
                iconName = 'Cancel';
                break;
              case '5000':
                iconName = 'ReturnToSession';
                break;
              case '9000':
                iconName = 'CompletedSolid';
                break;
              default:
                iconName = 'Refresh';
                break;
            }
        
            return (
              <div style={{ display: 'flex', flexDirection: 'row', alignItems: 'center' }}>
                <Icon iconName={iconName} />
                <span style={{ marginLeft: '8px', lineHeight: '24px' }}>{item.status}</span>
              </div>
            );
          },
        },
        
        
        { key: 'actionDate', name: 'Action Date', fieldName: 'actionDate', minWidth: 100, maxWidth: 150, isResizable: true ,
            onRender: (item) => {
               
                if (item.actionDate){
                    const formattedDate = format(new Date(item.actionDate), 'dd-MMM-yyyy');
                const formattedTime = format(new Date(item.actionDate), 'hh:mm a');
                return `${formattedDate} ${formattedTime}`;

                }
                return ''

                
              }
        } 
    ];

    return (
        <div style={{ overflowX: 'auto' }}>
            <DetailsList
                items={gridData} 
                columns={columns} 
                layoutMode={DetailsListLayoutMode.fixedColumns} 
                selectionMode={SelectionMode.none}
                isHeaderVisible={true} 
                styles={detailsListStyles}
            />
        </div>
    );
};

export default ApproverAndReviewerTableInViewForm;