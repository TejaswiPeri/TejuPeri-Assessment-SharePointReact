// Created by Teju Peri - *** 10-6-21 ***  Arlington County - MS Solution Architect Role

/* Not the complete example or the app/codebase for lack of time
But to give an idea of CRUD using SharePoint List/React
Dummy Code- Full Syntax check not done*/


import * as React from 'react';
import styles from './TejuSPReactCRUD.module.scss';
import { ITejuSPReactCRUDProps } from './ITejuSPReactCRUDProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
//If we want to use some Date fields or say leverage SharePoint PeoplePicker
import { DatePicker, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

import { WebPartContext } from "@microsoft/sp-webpart-base";




// This wil be in the File ITejuSPReactCRUDProps.ts 

export interface ITejuSPReactCRUD {
    description: string;
    context: WebPartContext;
    webURL: string;
}


/* This wil be in the File  TejuSPReactCRUDWebPart.ts file and import the component 
SharePointListReactCRUD.tsx in this file */
public render(): void {
    const element: React.ReactElement < ITejuSPReactCRUDProps > = React.createElement(
        CRUDReact,
        {
            description: this.properties.description,
            webURL: this.context.pageContext.web.absoluteUrl,
            context: this.context
        }
    );

    ReactDom.render(element, this.domElement);
}





//Interface Declaration
export interface IStates {
    Items: any;
    ID: any;
    name: any;
    status: any;
    HTML: any;
}



export default class SharePointListReactCRUD extends React.Component<ITejuSPReactCRUDProps, IStates> {
    //Everything empty set in the constructor
    constructor(props) {
        super(props);
        this.state = {
            Items: [],
            name: "",
            ID: 0,
            status: "",
            HTML: []

        };
    }

    //Get Items from SP List using the componentDidMount() method
    public async componentDidMount() {
        await this.fetchData();
    }

    public async fetchData() {

        let web = Web(this.props.webURL);
        const items: any[] = await web.lists.getByTitle("ToDoList").items.select("*", "name").get();
        console.log(items);
        this.setState({ Items: items });
        let html = await this.getHTML(items);
        this.setState({ HTML: html });
    }




    public async getHTML(items) {
        var tabledata = <table className={styles.table}>
            <thead>
                <tr>
                    <th>Task name</th>
                    <th>Task Status</th>
                </tr>
            </thead>
            <tbody>
                {items && items.map((item, i) => {
                    return [
                        <tr key={i} onClick={() => this.findData(item.ID)}>
                            <td>{item.name}</td>
                            <td>{item.status}</td>
                        </tr>
                    ];
                })}
            </tbody>

        </table>;
        return await tabledata;
    }
    /*
Adding an item to a hypothetical SharePoint Online/O365 list "To Do" in the SharePoint framework  using reactjs .
using PnP once the item gets created, we can call a setState() to make the form fields empty.
fetchData() method is then called, so that it will reload the data and we can see the updated items.*/


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
//setState()
//fetchData()
    private async DeleteDataSPList() {
       
    }
    private async UpdateDataSPList() {

    }

    public render(): React.ReactElement<ITejuSPReactCRUDProps> {
        return (
            <div >
                <h1>ReactJs - CRUD Operations </h1>
                {this.state.HTML}
                <div className={styles.btngroup}>
                    <div><PrimaryButton text="Create" onClick={() => this.SaveDataSPList()} /></div>
                    <div><PrimaryButton text="Update" onClick={() => this.UpdateDataSPList()} /></div>
                    <div><PrimaryButton text="Delete" onClick={() => this.DeleteDataSPList()} /></div>
                </div>
                <div>
                    <form>
                        <div>
                            // We can use SharePoint PeoplePicker  as well to get user name properties etc.
                            <Label>SharePoint User Name</Label>
                            <PeoplePicker
                            />
                        </div>
                        <div>
                            <Label> To Do List Task Name</Label>
                            <TextField value={this.state.name} multiline onChanged={(value) => this.onchange(value, "name")} />
                        </div>
                        <div>
                            <Label>Task Status</Label>
                            <TextField value={this.state.status} multiline onChanged={(value) => this.onchange(value, "status")} />
                        </div>

                    </form>
                </div>
            </div >
        );
    }







}

















