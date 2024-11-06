/* eslint-disable @typescript-eslint/no-explicit-any */

import * as React from 'react';
import { Checkbox, FontWeights, getTheme, Link, mergeStyleSets, MessageBar, MessageBarType, Modal, PrimaryButton } from '@fluentui/react';
import { IInputs } from '../generated/ManifestTypes';
import { Dialog, Icon, SelectionMode, TooltipHost } from '@fluentui/react';
import { useState } from 'react';
import * as XLSX from 'xlsx';
import { DetailsList, IColumn, IObjectWithKey } from '@fluentui/react';
import '../main.css';
import { PaginationComponent } from './PaginationComponent';
import { getData, postData, postDataQoute } from './DynamicService';
import { ConfirmationDialogComponent } from './ConfirmationDialogComponent';
import {Product} from '../Model';
export interface QuoteProductsComponentProps {
    label: string;
    onChanges: (newValue: string | undefined) => void;
    context: ComponentFramework.Context<IInputs>;
    quoteid:string;
    clientUrl:string;
}

const theme = getTheme();
const contentStyles = mergeStyleSets({
    container: {
        display: 'flex',
        flexFlow: 'column nowrap',
        alignItems: 'stretch',
    },
    header: [
        theme.fonts.xLargePlus,
        {
            flex: '1 1 auto',
            borderTop: `4px solid ${theme.palette.themePrimary}`,
            color: theme.palette.neutralPrimary,
            display: 'flex',
            alignItems: 'center',
            fontWeight: FontWeights.semibold,
            padding: '12px 12px 14px 24px',
        },
    ],
    heading: {
        color: theme.palette.neutralPrimary,
        fontWeight: FontWeights.semibold,
        fontSize: 'inherit',
        margin: '0',
    },
    body: {
        flex: '4 4 auto',
        padding: '0 24px 24px 24px',
        overflowY: 'hidden',
        selectors: {
            p: { margin: '14px 0' },
            'p:first-child': { marginTop: 0 },
            'p:last-child': { marginBottom: 0 },
        },
    },
});
const columns: IColumn[] = [
    // {
    //     key: 'column1', name: '', fieldName: 'isValidProduct', minWidth: 50, maxWidth: 200, isMultiline: false,
    //     onRender: (item: any) => (

    //         <Checkbox
    //             checked={item.isUploaded == 1} // You can manage checked state if needed
    //             disabled={item.isUploaded == 3} // Initially disabled
    //         />
    //     ),
    // },
    {
        key: 'column1', name: '', fieldName: 'isValidProduct', minWidth: 50, maxWidth: 200, isMultiline: false,
        onRender: (item: any) => (

            <TooltipHost
                content={item.message}
                calloutProps={{ gapSpace: 0 }}
            >
                <Icon iconName="Warning" styles={{ root: { color: item.isValidProduct ? (item.isUploaded==1?'green':'yellow') : 'red', fontSize: 24, cursor: 'pointer' } }} />
            </TooltipHost>
        ),
    },

    { key: 'column2', name: 'ProductNumber', fieldName: 'ProductNumber', minWidth: 100, maxWidth: 150, isMultiline: false },
    { key: 'column3', name: 'Uom', fieldName: 'Uom', minWidth: 50, maxWidth: 150, isMultiline: false },
    { key: 'column4', name: 'Price', fieldName: 'Price', minWidth: 50, maxWidth: 150, isMultiline: false },
    { key: 'column5', name: 'Quantity', fieldName: 'Quantity', minWidth: 50, maxWidth: 150, isMultiline: false },
];

export const QuoteProductsComponent = React.memo((props: QuoteProductsComponentProps) => {
    const [isDialogOpen, setIsDialogOpen] = useState(false);
    const openDialog = (): any => setIsDialogOpen(true);
    const closeDialog = (): any => setIsDialogOpen(false);
    const [isConfirmDialogOpen, setIsConfirmDialogOpen] = useState(false);
    const openConfirmDialog = (): any => setIsConfirmDialogOpen(true);
    const closeConfirmDialog = (): any => setIsConfirmDialogOpen(false);
    const [products, setProducts] = React.useState<Product[]>([]);
    const [isProducts, setIsProducts] = useState(false);
    const [currentPage, setCurrentPage] = useState(1);
    const [pageSize, setPageSize] = useState(5);
    const [totalPages, setTotalPages] = useState(0);
    const [paginatedItems, setPaginatedItems] = React.useState<Product[]>([]);
    const [isUploading, setIsUploading] = React.useState(false);
    const [isExcel, setIsExcel] = useState(true);
    const [fileName, setFileName] = useState('');
    const[isUploaded,setIsUploaded] = useState(false);

  React.useEffect(()=>{
    setIsUploaded(false);
    //uploadedData()
  },[]);

    const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        const isExc = isExcelFile(file);
        if (isExc) {
            setIsExcel(true);
            setFileName(file?.name.toString()||'');
            if (file) {
                const reader = new FileReader();
                reader.onload = (e) => {
                    const data = new Uint8Array(e.target?.result as ArrayBuffer);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonProducts: Product[] = XLSX.utils.sheet_to_json(worksheet);
                    validateProducts(jsonProducts);
                };
                reader.readAsArrayBuffer(file);
            }
        } else {
            setFileName(file?.name.toString()||'');
            setIsExcel(false);
        }
    };

    const isExcelFile = (file: any) => {
        const validExtensions = ['.xls', '.xlsx', '.xlsm'];
        const fileExtension = file.name.split('.').pop().toLowerCase();
        return validExtensions.includes(`.${fileExtension}`);
    };

    const validateProducts = (jsonProducts: Product[]) => {
        const allProducts: Product[] = [];
        //const invalid = [];

        jsonProducts.forEach((product) => {
            const isValid = product.Price >= 0 && product.Uom && product.ProductNumber;
            product.isUploaded = 3;
            if (isValid) {
                product.isValidProduct = true;
                product.message = 'Good to Go';
                allProducts.push(product);
            } else {
                product.isValidProduct = false;
                product.message = (product.Price < 0) ? 'Price is Less Than 0' : ((product.Uom) ? 'Product Number is not Present' : 'Product UOM is not Present');
                allProducts.push(product);
            }
        });
        const total = Math.ceil(allProducts.length / pageSize);
        setTotalPages(total);
        setProducts(allProducts);
        const paginated = allProducts.slice((currentPage - 1) * pageSize, currentPage * pageSize);
        setPaginatedItems(paginated);
        setIsProducts(true);
    };

    const saveToQuote = () => {

    }
    const handleButtonClick = () => {
        const fileInput: any = document.getElementById('fileUpload');
        fileInput.click();  // Trigger the click on the hidden file input
    };

    const closeAndClear = () => {
        closeDialog();
        setProducts([]);
        setCurrentPage(1);
        setPaginatedItems([]);
        setTotalPages(0);
        setFileName('');
        setIsProducts(false);
        setIsUploaded(false);
        window.location.reload();
    }

    const handleCurrentPage = (page: any) => {
        setCurrentPage(page);
        const total = Math.ceil(products.length / pageSize);
        setTotalPages(total);
        const paginated = products.slice((page - 1) * pageSize, page * pageSize);
        setPaginatedItems(paginated);
    }
    const handlePageSizeChange = (pageSize: any) => {
        setPageSize(pageSize);
        setCurrentPage(1);
        const total = Math.ceil(products.length / pageSize);
        setTotalPages(total);
        const paginated = products.slice((currentPage - 1) * pageSize, currentPage * pageSize);
        setPaginatedItems(paginated);
    }

    // const saveQuoteData = async () => {
    //     setIsUploading(true);
    //     const productResponse = await getData('https://org01fafc2a.api.crm.dynamics.com', 'products?$select=productid,productnumber,name');
    //     const uoms = await getData('https://org01fafc2a.api.crm.dynamics.com', 'uoms?$select=uomid,name');

    //     const quoteName = 'Quote Name is :' + ' LapTop Products';
    //     const quoteData: any = { name: quoteName };
    //     quoteData["customerid_account@odata.bind"] = '/accounts(88019079-e334-46c0-b2da-27d8a73c51dc)'; // Customer
    //     quoteData["pricelevelid@odata.bind"] = '/pricelevels(07897316-500b-ea11-a813-000d3a1b1808)';
    //     const quoteid = await postDataQoute('https://org01fafc2a.api.crm.dynamics.com', 'quotes', quoteData);

    //     for (const data of products) {
    //         if (productResponse.value) {
    //             const product = productResponse.value.find((d: any) => d.productnumber == data.ProductNumber);
    //             const uom = uoms.value.find((d: any) => d.name == data.Uom);
    //             const quoteDetailData: any = {};
    //             quoteDetailData.priceperunit = data.Price;
    //             quoteDetailData["quoteid@odata.bind"] = `/quotes(${quoteid})`;
    //             quoteDetailData.quantity = data.Quantity;
    //             quoteDetailData["uomid@odata.bind"] = `/uoms(${uom.uomid})`;
    //             quoteDetailData["productid@odata.bind"] = `/products(${product.productid})`;

    //             const detailid = await postDataQoute('https://org01fafc2a.api.crm.dynamics.com', 'quotedetails', quoteDetailData);
    //             console.log(detailid);
    //         }

    //     }
    //     setIsUploading(false);
    // }
    const uploadedData = (data: any) => {
        const newArray  = [...products,data];
        setProducts(newArray);
        const total = Math.ceil(data.length / pageSize);
        setTotalPages(total);
        const paginated = data.slice((currentPage - 1) * pageSize, currentPage * pageSize);
        setPaginatedItems(paginated);
        setIsUploaded(true);
    }


    return (
        <>
            <button onClick={openDialog}> import Products</button>

            <Modal
                isOpen={isDialogOpen}
                onDismiss={closeDialog}
                isBlocking={false}
                containerClassName={contentStyles.container}
                // modalProps={modalProps}
                styles={{
                    main: {
                        selectors: {
                            ['@media (min-width: 480px)']: {
                                // minWidth: 450,
                                // maxWidth: '1500px',
                                // width: '1300px',
                                padding: '30px',
                                // height: '600px'
                            }
                        }
                    }
                }}
            >
                <div className='main-container'>
                    <div className='header_text' style={{marginBottom:isExcel?'29px':'0px'}}>
                        <div>
                            <h2>Excel Export</h2>
                            <MessageBar>
                                Pease Choose Only Excel File for Adding Quotes
                            </MessageBar>
                        </div>
                        <Icon iconName="Cancel" onClick={closeAndClear} className='cut-icon' style={{ cursor: 'pointer', fontSize: '24px', color: 'red' }} />
                    </div>
                    {!isExcel ?
                        <MessageBar
                            messageBarType={MessageBarType.error}
                            isMultiline={false}
                        >
                            Please Select valid Excel File Thanks.

                        </MessageBar> : null}
                    <input
                        type="file"
                        id="fileUpload"
                        accept=".xls,.xlsx"
                        onChange={handleFileChange}
                        style={{ display: 'none' }}  /* Hide the file input */
                    />
                    <div className='dflex-column'>
                        <PrimaryButton onClick={handleButtonClick}>Choose File</PrimaryButton>
                        <p style={{ display: fileName == '' ? 'none' : 'block' }}>{fileName}</p>
                    </div>
                    <ConfirmationDialogComponent isopen={isConfirmDialogOpen} onClose={closeConfirmDialog} uploadedData={uploadedData} products={products} quoteid={props.quoteid} clientUrl={props.clientUrl}/>
                    {isExcel ?
                        <div>
                            <div style={{ display: isProducts ? 'block' : 'none' }}>
                                <DetailsList
                                    items={paginatedItems}  
                                    columns={columns}
                                    setKey="set"
                                    layoutMode={0}
                                    compact={true}
                                    selectionMode={SelectionMode.none}
                                />
                                <PaginationComponent
                                    currentPage={currentPage}
                                    totalPages={totalPages}
                                    pageSize={pageSize}
                                    onPageChange={handleCurrentPage}
                                    onPageSizeChange={handlePageSizeChange}
                                />

                            </div>

                            <div className='mt-86' style={{ display: isProducts ? 'block' : 'none' }}>
                                <PrimaryButton className='fright' onClick={openConfirmDialog} disabled={isUploaded}>Add Products To Quote</PrimaryButton>
                            </div>
                        </div> : null}
                </div>
            </Modal>


        </>
    );
});