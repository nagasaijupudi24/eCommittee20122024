/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from 'react';

import {
  DetailsList,
  Selection,
  IColumn,

 
  IDragDropEvents,
  IDragDropContext,
  SelectionMode,
  IDetailsFooterProps,
} from '@fluentui/react/lib/DetailsList';

import { getTheme, IconButton, mergeStyles } from '@fluentui/react';
import { IExampleItem } from '@fluentui/example-data';

const theme = getTheme();

const dragEnterClass = mergeStyles({
  backgroundColor: theme.palette.neutralLight,
});


interface IDetailsListDragDropExampleState {
  items: any;
  columns: IColumn[];
 
}


export class DetailsListDragDropExample extends React.Component<any, IDetailsListDragDropExampleState> {
  private _selection: Selection;
  private _dragDropEvents: IDragDropEvents;
  private _draggedItem: any[] | undefined;
  private _draggedIndex: number;
  private _columns:any =[
    {
      key: 'dragHandle',
      name: '',
      fieldName: 'dragHandle',
      minWidth: 50,
      maxWidth: 50,
      isResizable: false,
      onRender: (item: any) => (
        <div >  <IconButton
        iconProps={{ iconName: 'GlobalNavButton' }} 
        title="Menu"
        ariaLabel="Menu"
        
    /></div>
      
      ),
    },
    {
      key: 'serialNo',
      name: 'S.No',
      
      minWidth: 50,
      maxWidth: 80,
      isResizable: false,
      onRender: (_item: any, _index?: number) => (
        <div style={{ marginTop: '8px' }}>{(_index !== undefined ? _index : 0) + 1}</div>
      ),
    },
    {
      key: 'text',
      name:this.props.type, 
      fieldName: 'text',
      minWidth: 100,
      maxWidth: 295,
      isResizable: true,
      onRender: (item: any) => (
        <div style={{ marginTop: '8px' }}>{item.text}</div> 
      ),
    },
    {
      key: 'srNo',
      name: 'SR No',
      fieldName: 'srNo',
      minWidth: 100,
      maxWidth: 295,
      isResizable: true,
      onRender: (item: any) => (
        <div style={{ marginTop: '8px' }}>{item.srNo}</div> 
      ),
    },
    {
      key: 'optionalText',
      name: 'Designation',
      fieldName: 'optionalText',
      minWidth: 100,
      maxWidth: 295,
      isResizable: true,
      onRender: (item: any) => (
        <div style={{ marginTop: '8px' }}>{item.optionalText}</div> 
      ),
    },
    {
      key: 'actions',
      name: 'Actions',
      fieldName: 'actions',
      minWidth: 50,
      maxWidth: 80,
      isResizable: false,
      onRender: (_item: any) => (
        <IconButton
          iconProps={{ iconName: 'Delete' }} 
          title="Delete"
          ariaLabel="Delete"
          onClick={()=>{
           
            this._remove(_item)
          }} 
        />
      ),
    },
  ];


  private _remove = (dataItem:any) => {
    this.props.removeDataFromGrid(dataItem,this.props.type)
    
  };
  constructor(props: any) {
    super(props);

    this._selection = new Selection();
    this._dragDropEvents = this._getDragDropEvents();
    this._draggedIndex = -1;
   

    this.state = {
      items:this.props.data,
      columns:this._columns,
     
    };
    
  }

  componentDidMount() {
    try {
       
        if (this.props.data !== this.state.items) {
            this.setState({ items: this.props.data });
        }
    } catch (error) {
        console.error("Error in componentDidMount:", error);
    }
}


  public render(): JSX.Element {
    const {  columns,items } = this.state;
  

    return (
      <div>
        <div 
       
        >
    
        </div>
      
        <DetailsList
          setKey="items"
          items={items}
          columns={columns}
          selection={this._selection}
          selectionMode={SelectionMode.none}
          selectionPreservedOnEmptyClick={true}
          dragDropEvents={this._dragDropEvents}
          onRenderDetailsFooter={(props: IDetailsFooterProps) => {
            if (this.state.items.length === 0) {
              return (
                <div style={{ textAlign: 'center', padding: '20px', color: 'gray' }}>
                  No records available
                </div>
              );
            }
            return null;
          }}
        />
      
      </div>
    );
  }




  private _getDragDropEvents(): IDragDropEvents {
    return {
      canDrop: (dropContext?: IDragDropContext, dragContext?: IDragDropContext) => {
        return true;
      },
      canDrag: (item?: any) => {
        return true;
      },
      onDragEnter: (item?: any, event?: DragEvent) => {
      
        return dragEnterClass;
      },
      onDragLeave: (item?: any, event?: DragEvent) => {
        return;
      },
      onDrop: (item?: any, event?: DragEvent) => {
        if (this._draggedItem) {
          this._insertBeforeItem(item);
        }
      },
      onDragStart: (item?: any, itemIndex?: number, selectedItems?: any[], event?: MouseEvent) => {
        this._draggedItem = item;
        this._draggedIndex = itemIndex!;
      },
      onDragEnd: (item?: any, event?: DragEvent) => {
        this._draggedItem = undefined;
        this._draggedIndex = -1;
      },
    };
  }


private _insertBeforeItem(item: IExampleItem): void {
  const draggedItems = this._selection.isIndexSelected(this._draggedIndex)
    ? (this._selection.getSelection() as IExampleItem[])
    : [this._draggedItem!];

  const insertIndex = this.state.items.indexOf(item);
  const items = this.state.items.filter((itm: any) => draggedItems.indexOf(itm) === -1);

  items.splice(insertIndex, 0, ...draggedItems);

  this.setState({ items:items });
 
  this.props.reOrderData(items,this.props.type);
}
}
