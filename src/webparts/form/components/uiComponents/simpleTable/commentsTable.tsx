/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/ban-ts-comment */
import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';

const CommentsLogTable = (props: any) => {
    const gridData = props.data;

    // Define the columns for the DetailsList
    const columnsNew: IColumn[] = [
        { key: 'pageNum', name: 'Page#', fieldName: 'pageNum', minWidth:80, maxWidth: 265, isResizable: true },
        { key: 'page', name: 'Doc Reference', fieldName: 'page', minWidth: 80, maxWidth: 265, isResizable: true },
        { key: 'comment', name: 'Comments', fieldName: 'comment', minWidth: 80, maxWidth:265, isResizable: true, isMultiline: true },
        { key: 'commentedBy', name: 'Comment By', fieldName: 'commentedBy', minWidth: 80, maxWidth: 265
            , isResizable: true }
    ];


    const columnsView: IColumn[] = [
        { key: 'pageNum', name: 'Page#', fieldName: 'pageNum', minWidth:80, maxWidth: 150, isResizable: true },
        { key: 'page', name: 'Doc Reference', fieldName: 'page', minWidth: 80, maxWidth: 150, isResizable: true },
        { key: 'comment', name: 'Comments', fieldName: 'comment', minWidth: 80, maxWidth: 250, isResizable: true, isMultiline: true },
        { key: 'commentedBy', name: 'Comment By', fieldName: 'commentedBy', minWidth: 80, maxWidth: 150, isResizable: true }
    ];

    switch (props.type) {
        case "generalComments":
            return <div>{" "}</div>;
        case "commentsLog":
            return (
                <div style={{ overflowX: 'auto' }}>
                    <DetailsList
                        items={gridData} // Data for the table
                        columns={props.formType === 'new'?columnsNew:columnsView} // Column definitions
                        layoutMode={DetailsListLayoutMode.fixedColumns} // Fixed column layout
                        selectionMode={SelectionMode.none} // Disable row selection
                        isHeaderVisible={true} // Show header
                    />
                </div>
            );
        default:
            return <div>{" "}</div>;
    }
};

export default CommentsLogTable;
