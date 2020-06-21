export default interface IMailAttachment {
    '@odata.mediaContentType': string;
    '@odata.type': string;
    contentBytes: string;
    contentId: string;
    contentLocation?: string;
    contentType: string;
    id: string;
    isInline: boolean;
    lastModifiedDateTime: string;
    name: string;
    size: number;
}