/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import '../main.css';
import { Icon } from '@fluentui/react';
export const PaginationComponent: React.FC<any> = ({ currentPage, totalPages, pageSize, onPageChange, onPageSizeChange }) => {
    const handlePageSizeChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
        onPageSizeChange(Number(event.target.value));
       // onPageChange(1); // Reset to first page
    };

    const handlePrev = () => {
        if (currentPage > 1) {
            onPageChange(currentPage - 1);
        }
    };

    const handleNext = () => {
        if (currentPage < totalPages) {
            onPageChange(currentPage + 1);
        }
    };

    return (
        <>
            <div className='paginated-top'>
                <p>Rows per page</p>
                <select value={pageSize} onChange={handlePageSizeChange}>
                    <option value={5}>5</option>
                    <option value={10}>10</option>
                    <option value={20}>20</option>
                    <option value={50}>50</option>
                </select>
                <span>Page {currentPage} of {totalPages}</span>
                {/* <button onClick={handlePrev} disabled={currentPage === 1}>Previous</button> */}
                <Icon
                    iconName="NavigateBack"
                    className='icn'
                    style={{ display: currentPage != 1 ? 'block' : 'none' }}
                    onClick={handlePrev}
                />
                {/* <button onClick={handleNext} disabled={currentPage === totalPages}>Next</button> */}
                <Icon
                    iconName="NavigateForward"
                    style={{ display: currentPage != totalPages ? 'block' : 'none' }}
                    onClick={handleNext}
                    className='icn'
                />
            </div>
        </>
    );

}