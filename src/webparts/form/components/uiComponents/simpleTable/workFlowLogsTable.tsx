/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/ban-ts-comment */
import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { format } from 'date-fns';

const WorkFlowLogsTable = (props: any) => {
    const { data } = props;

  
    const formatDateTime = (date: string | number | Date) => {
        const formattedDate = format(new Date(date), 'dd-MMM-yyyy');
        const formattedTime = format(new Date(date), 'hh:mm a');
        return `${formattedDate} ${formattedTime}`;
    };

    
    const columns: IColumn[] = [
        { key: 'action', name: 'Action', fieldName: 'action', minWidth: 200, maxWidth: 250, isResizable: true },
        { key: 'actionBy', name: 'Action By', fieldName: 'actionBy', minWidth: 80, maxWidth: 150, isResizable: true },
        {
            key: 'createdDate',
            name: 'Action Date',
            fieldName: 'createdDate',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            onRender: (item: any) => (
                <span>{formatDateTime(item.createdDate)}</span>
            ),
        },
    ];

    return (
        <div style={{ overflowX: 'auto' }}>
            <DetailsList
                items={data} 
                columns={columns}
                layoutMode={DetailsListLayoutMode.fixedColumns} 
                selectionMode={SelectionMode.none} 
                isHeaderVisible={true} 
            />
        </div>
    );
};

export default WorkFlowLogsTable;
