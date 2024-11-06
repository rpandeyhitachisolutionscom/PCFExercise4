/* eslint-disable @typescript-eslint/no-explicit-any */
import * as ReactDOM from "react-dom";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { QuoteProductsComponent } from './components/QuoteProductsComponent';
export class ImportQuoteProducts implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    notifyOutputChanged: () => void;
    rootContainer: HTMLDivElement;
    selectedValue: string | null;
    context: ComponentFramework.Context<IInputs>;

    constructor() {

    }
    onChange = (newValue: string | undefined): void => {
        this.context.parameters.value.raw = newValue || '';
        this.selectedValue = newValue || '';
        this.notifyOutputChanged();
    };

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.rootContainer = container;
        this.context = context;
        const value = context.parameters.value;
        const currentUrl = window.location.href;

        // Extract the ID from the URL
        const id = this.getIdFromUrl(currentUrl);
        console.log(id);
        const clientUrl = (
            window as any
          ).parent.Xrm.Utility.getGlobalContext().getClientUrl();

        ReactDOM.render(
            React.createElement(QuoteProductsComponent, {
                label: value.raw || '',
                onChanges: this.onChange,
                context: this.context,
                quoteid: id || '',
                clientUrl:clientUrl
            }),
            this.rootContainer,
        );
    }


    private getIdFromUrl(url: string): string | undefined {
        try {
            const parsedUrl = new URL(url);
            const params = new URLSearchParams(parsedUrl.search);
            return params.get("id") || undefined; // Returns the ID or undefined if not found
        } catch (error) {
            console.error('Error parsing URL:', error);
            return undefined;
        }
    }


    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Add code to update control view
    }

    public getOutputs(): IOutputs {
        return { value: this.selectedValue } as IOutputs;
    }


    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this.rootContainer);
    }
}
