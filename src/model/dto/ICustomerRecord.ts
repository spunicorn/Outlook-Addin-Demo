

import { IProduct } from "./IProduct";
import { ICustomer } from "./ICustomer";


export interface ICustomerRecord{
    Title:string;
    ProductInfo:IProduct
    CustomerInfo:ICustomer
    DatePurchased:Date;
    DateExpires:Date;
}