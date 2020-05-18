import { ICustomerRecord } from "../model/dto/ICustomerRecord";

export interface ICustomerRepo{
    GetCustomersByEmail(email:string) : Promise<ICustomerRecord[]>
    GetAllCustomers() : Promise<ICustomerRecord[]>
}

export class MockCustomerRepo implements ICustomerRepo{
    constructor(mockData:ICustomerRecord[]){
        this.MockData = mockData;
    }
    MockData:ICustomerRecord[];
    protected FilterByEmail(  custRecord:ICustomerRecord,email:string):ICustomerRecord[]{
        
        return new Array<ICustomerRecord>();
    }
    GetCustomersByEmail(email:string): Promise<ICustomerRecord[]> {
        return new Promise<ICustomerRecord[]>((resolve,reject) =>{
            setTimeout(function(){
                resolve( this.MockData.filter(customer =>{
                    return customer.Email.toLocaleLowerCase() == email.toLocaleLowerCase();
                }));
            }, 1000);
            
            
        })
        
    }
    GetAllCustomers(): Promise<ICustomerRecord[]> {
        throw new Error("Method not implemented.");
    }

}