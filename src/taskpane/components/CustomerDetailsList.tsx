import * as React from 'react';
import { Announced } from 'office-ui-fabric-react/lib/Announced';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { ICustomerRecord } from '../../model/dto/ICustomerRecord';
import { SPClient } from '../../dal/SPClient';
import { Button, ButtonType } from 'office-ui-fabric-react';

const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
});

const textFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };

export interface IDetailsListBasicExampleItem {
  key: number;
  name: string;
  value: number;
}

export interface IDetailsListBasicExampleState {
  customerRelatedProductsFiltered:ICustomerRecord[];
  selectionDetails: string;
  customerRelatedProducts:ICustomerRecord[];
  customerDetailsLoading:false;
}
export interface IPropseroni{
    spClient:SPClient
}
export default class CustomerDetailsList extends React.Component<IPropseroni, IDetailsListBasicExampleState> {
  private _selection: Selection;
  
  private _columns: IColumn[];
//   protected spClient: SPClient = new SPClient();

  constructor(props :IPropseroni) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    // Populate with items for demos.
    

    this._columns = [
      { key: 'column1', name: 'Customer', fieldName: 'name', 
      onRender:(item,index,column)=>{return item.CustomerInfo.Title} 
      ,minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'Product', fieldName: 'value',
      onRender:(item,index,column)=>{return item.ProductInfo.Title}
      , minWidth: 100, maxWidth: 200, isResizable: true },
    ];

    this.state = {
      customerRelatedProductsFiltered:[],
      customerRelatedProducts:[],
      customerDetailsLoading:false,
      selectionDetails: this._getSelectionDetails(),
    };
  }
  searchSP = async ()=>{
    let fromEmail = Office.context.mailbox.item.from;
    let customerRecords = await this.props.spClient.getProductsRelatedCustomer(fromEmail.emailAddress);
    this.setState({
      customerRelatedProducts : customerRecords,
      customerRelatedProductsFiltered: customerRecords
    })
  }
  public render(): JSX.Element {
    const { customerRelatedProducts, selectionDetails,customerRelatedProductsFiltered } = this.state;

    return (
      <Fabric>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.searchSP.bind(this)}
          >
            Get customer records
          </Button>

        <div className={exampleChildClass}>{selectionDetails}</div>
        <Announced message={selectionDetails} />
        <TextField
          className={exampleChildClass}
          label="Filter by product name:"
          onChange={this._onFilter}
          styles={textFieldStyles}
        />
        <Announced message={`Number of items after filter applied: ${customerRelatedProducts.length}.`} />
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={customerRelatedProductsFiltered}
            columns={this._columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
            onItemInvoked={this._onItemInvoked}
          />
        </MarqueeSelection>
      </Fabric>
    );
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as ICustomerRecord).ProductInfo.Title;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
        customerRelatedProductsFiltered : text ? this.state.customerRelatedProducts.filter(i => i.ProductInfo.Title.toLowerCase().indexOf(text) > -1) : this.state.customerRelatedProductsFiltered,
    });
  };

  private _onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
    alert(`Item invoked: ${item.name}`);
  };
}