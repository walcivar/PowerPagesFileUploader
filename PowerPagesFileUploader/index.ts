import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { FileUploader } from "./FileUploader";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";

export class PowerPagesFileUploader implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _root: Root | null = null;

    constructor() { }

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = document.createElement("div");
        this._container.style.width = "100%";
        this._container.style.height = "100%";
        container.appendChild(this._container);
        context.mode.trackContainerResize(true);

        this._root = createRoot(this._container);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;

        const props = {
            context: this._context,
            notifyOutputChanged: this._notifyOutputChanged,
            height: this._context.parameters.Height.raw || "200px",
            width: this._context.parameters.Width.raw || "100%",
            blobAccountName: this._context.parameters.BlobAccountName.raw || "",
            blobContainerName: this._context.parameters.BlobContainerName.raw || "",
            sasTokenEnvironmentVariable: this._context.parameters.SasTokenEnvironmentVariable.raw || "",
            powerPagesUrl: this._context.parameters.PowerPagesUrl.raw || ""
        };

        if (this._root) {
            this._root.render(React.createElement(FileUploader, props));
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        if (this._root) {
            this._root.unmount();
        }
    }
}
