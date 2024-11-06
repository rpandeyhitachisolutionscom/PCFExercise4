/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import '../main.css';
import { Dialog, DialogFooter, DialogType, Icon, Label, PrimaryButton, Spinner, SpinnerSize } from '@fluentui/react';
import { getData, postData, postDataQoute } from './DynamicService';

const dialogContentProps = {
    type: DialogType.normal,
    title: 'Missing Subject',
    closeButtonAriaLabel: 'Close',
    subText: 'Do you want to upload quotes into the system ?',
};
export const ConfirmationDialogComponent: React.FC<any> = ({ isopen, onClose, uploadedData, products, quoteid,clientUrl }) => {
    const [isUploading, setIsUploading] = React.useState(false);
    const [data, setData] = React.useState<any>([]);
    const [isUploaded, setIsUploaded] = React.useState(false);
    const confirToUpload = () => {

    }
    const notConfirmToUpload = () => {
        onClose();
    }

    const saveQuoteData = async () => {
        setIsUploading(true);
        const productQuery = 'products?$select=productid,productnumber,name';
        const productResponse = await getData(clientUrl, productQuery);
        const uomQuery = 'uoms?$select=uomid,name';
        const uoms = await getData(clientUrl, uomQuery);

        for (const data of products) {
            if (productResponse.value) {
                const product = productResponse.value.find((d: any) => d.productnumber == data.ProductNumber);
                const uom = uoms.value.find((d: any) => d.name == data.Uom);
                const quoteDetailData: any = {};
                if (!product) { data.isUploaded = 0; data.isValidProduct = false; data.message = 'Product is not matched '; continue; }
                if (!uom) { data.isUploaded = 0; data.isValidProduct = false; data.message = 'UOM is not matched '; continue; }
                quoteDetailData.ispriceoverridden = true;
                quoteDetailData.priceperunit = data.Price;
                quoteDetailData["quoteid@odata.bind"] = `/quotes(${quoteid})`;
                quoteDetailData.quantity = data.Quantity;
                quoteDetailData["uomid@odata.bind"] = `/uoms(${uom.uomid})`;
                quoteDetailData["productid@odata.bind"] = `/products(${product.productid})`;

                const detailid = await postDataQoute(clientUrl, 'quotedetails', quoteDetailData);
                if (!detailid) { data.isUploaded = 0; data.isValidProduct = false; data.message = 'Internal Error '; continue; }
                if (detailid) { data.isUploaded = 1; data.message = 'Uploaded Successfull '; continue; }
                console.log(detailid);
            }

        }
        // alert("Uploaded Successfully");
        setIsUploading(false);
        uploadedData(products);
        setIsUploaded(true);
        //  onClose();
    }

    const BeforeUpload = () => {
        return (
            <>
                {isUploading ?
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                        <Spinner label=" " size={SpinnerSize.large} styles={{ root: { width: 60, height: 60 } }} />
                        <Label style={{ marginLeft: 10, fontSize: '22px' }}>Saving Products to Quote</Label>
                    </div> :
                    <div>
                        <div>
                            <DialogFooter>
                                <PrimaryButton onClick={notConfirmToUpload} className='notcorrect'>NO</PrimaryButton>
                                <PrimaryButton onClick={saveQuoteData} className='correct'>YES</PrimaryButton>
                            </DialogFooter>
                        </div>
                    </div>
                }
            </>
        );
    }
    const AfterUpload = () => {
        return (

            <>
                <div>
                    <div> Data Uploaded SuccessFully </div>
                    <DialogFooter>
                        <PrimaryButton onClick={()=>onClose()} className='correct'>OK</PrimaryButton>
                    </DialogFooter>
                </div>
            </>
        );
    }

    return (
        <>
            <Dialog
                hidden={!isopen}
                onDismiss={onClose}
                dialogContentProps={dialogContentProps}
            >
                {isUploaded ? <AfterUpload /> : <BeforeUpload />}
            </Dialog>
        </>
    );
}