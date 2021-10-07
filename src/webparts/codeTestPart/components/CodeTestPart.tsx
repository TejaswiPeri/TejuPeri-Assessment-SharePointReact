import * as React from 'react';
import styles from './CodeTestPart.module.scss';
import { ICodeTestPartProps } from './ICodeTestPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';

// Modified by Teju Peri - *** 10-6-21 ***  Arlington County - MS Solution Architect Role


/*For simplicity sake I am just changing per the requirement. 
Please see my solution/readme for a more extensive solution/extension for this*/

export interface IDetailsListBasicExampleItem {
  key: number;
  name: string;
  status: string;
}

export interface IDetailsListBasicExampleState {
    items: IDetailsListBasicExampleItem[];
    selectionDetails: string; /* Added selection Details. We can press enter/focus on item row etc.and show alerts
    of which Item form the ToDo list has been selected*/

}

export default class CodeTestPart extends React.Component<ICodeTestPartProps, IDetailsListBasicExampleState> {
  private _selection: Selection;
  private _allItems: IDetailsListBasicExampleItem[];
  private _columns: IColumn[];

    /* I have alos tried to show how to use the Property '_selection',
    gave it an initializer and assigned it in the constructor.*/
    
  constructor(props: ICodeTestPartProps) {
      super(props);
      //Here I am filling with 3 dummy values in the Task List
      this._allItems = [];
      for (let i = 0; i < 3; i++) {
          this._allItems.push({
              key: i,
              name: 'ItemNumber ' + i,
              status: 'Pending',
          });
      }


      this._selection = new Selection({
          onSelectionChanged: () => this.setState({ selectionDetails: this.getSelectionDetails() }),
      });

    this._columns = [
      { key: 'column1', name: 'Task Name', fieldName: 'name', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'Status', fieldName: 'status', minWidth: 100, maxWidth: 200, isResizable: true },
    ];

    this.state = {
        items: this._allItems,
        selectionDetails: this.getSelectionDetails(),
    };
  }
  public render(): React.ReactElement<ICodeTestPartProps> {
      const { items, selectionDetails } = this.state;
      /* In the code below we can use focus on a item row, click etc. and excute a function like
      "onItemInvoked" to show an alert etc. We can use an Announced/Text to show the alert
      */
    return (
      <div className={styles.codeTestPart}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <DetailsList
                items={items}
                columns={this._columns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selection={this._selection}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }

  //Some kind of functions here to add the selection Details Logic to our app. 

  private getSelectionDetails(): string {
      const selectionCount = this._selection.getSelectedCount();

      switch (selectionCount) {
          case 0:
              return 'No task items have been selected';
          case 1:
              return 'One task item selected: ' + (this._selection.getSelection()[0] as IDetailsListBasicExampleItem).name;
          default:
              return `${selectionCount} items selected`;
      }
  }
  private onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
      alert(`An Task Item has ben selected: ${item.name}`);
  };
}

/*
Adding an item to a hypothetical SharePoint Online/O365 list "To Do" in the SharePoint framework  using reactjs .
*/
private async SaveDataSPList() {
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("ToDoList").items.add({

        name: this.state.name,
        status: this.state.status,

    }).then(i => {
        console.log(i);
    });
    alert("Created a new item in the To DO List Successfully");
    this.setState({ name: "", status: "" });
    this.fetchData();
}

//Please Check SharePointListReactCRUD.tsx for the complete code

