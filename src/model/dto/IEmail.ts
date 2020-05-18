export interface IEmail{
    id: string;
    receivedDateTime: string;
    sender: {
        emailAddress:{
            name: string;
            address: string;
        }
    };
    subject: string;
    bodyPreview: string;
}