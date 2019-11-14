import { IInputs, IOutputs } from "./generated/ManifestTypes";

// Import & attach jQuery to Window
import * as $ from "jquery";
declare var window: any;
window.$ = window.jQuery = $;

import * as DynamicsWebApi from "dynamics-web-api";
import { lookup } from "dns";
import * as convert from "xml-js";
//import "colresizable";

export class CustomFetchGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    /** HTML elements */
    private _grid: HTMLElement;

    /** Properties */
    private _primaryEntityName: string;
    private _primaryEntityNamePlural: string;
    private _fetchXML: string;

    private _jsonFetch: any;
    private _entitiesInFetch: Array<{ EntityLogicalName: string, EntityDisplayName: string, Alias: string }>;

    private _primaryFieldName: string;

    private _headers: Array<string>;
    private _headerDisplayNames: Array<{ LogicalName: string, DisplayName: string }>;
    private _itemsPerPage: number;

    private _currentPageNumber: number;
    private _totalNumberOfRecords: number;

    private _lock: boolean = false;

    private _webApi: DynamicsWebApi;

    /** Events */
    private _firstButton_Click: EventListenerOrEventListenerObject;
    private _previousButton_Click: EventListenerOrEventListenerObject;
    private _nextButton_Click: EventListenerOrEventListenerObject;

    private _th_Click: EventListenerOrEventListenerObject;
    private _link_Click: EventListenerOrEventListenerObject;

    /** General */
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _container: HTMLDivElement;

	/**
	 * Empty constructor.
	 */
    constructor() {

    }

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
        // Assigning environment variables.
        this.initVars(context, notifyOutputChanged, container);
        //this.loadPageFirst();
        this.initEventHandlers();
    }

    private initVars(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, container: HTMLDivElement): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;

        this._webApi = new DynamicsWebApi({ webApiVersion: '9.0' });
        this._headers = new Array<string>();
        this._headerDisplayNames = new Array<{ LogicalName: string, DisplayName: string }>();
        this._itemsPerPage = 5;
        this._currentPageNumber = 1;

        var globalContext = Xrm.Utility.getGlobalContext();
        var appUrl = globalContext.getCurrentAppUrl();

        this._container.style.overflow = "auto";
    }

    private initEventHandlers(): void {
        this._firstButton_Click = this.firstButton_Click.bind(this);
        this._previousButton_Click = this.previousButton_Click.bind(this);
        this._nextButton_Click = this.nextButton_Click.bind(this);
        this._th_Click = this.th_Click.bind(this);
        this._link_Click = this.link_Click.bind(this);
    }

    private firstButton_Click(evt: Event): void {
        this._currentPageNumber = 1;

        this.loadPage();
    }

    private previousButton_Click(evt: Event): void {
        this._currentPageNumber--;

        this.loadPage();
    }

    private nextButton_Click(evt: Event): void {
        this._currentPageNumber++;

        this.loadPage();
    }

    private th_Click(evt: Event): void {
        this._currentPageNumber = 1;
        let target = evt.currentTarget;

        // @ts-ignore
        var currentOrder = target.dataset.order;
        // @ts-ignore
        var headerName = target.dataset.logicalName;

        $(".uci-order-icon").removeClass("wj-glyph-down").removeClass("wj-glyph-up");
        $(".uci-thead-th").data("order", "none");

        this.OrderRecords(currentOrder, target, headerName);

        this.loadPage();
    }

    private OrderRecords(currentOrder: any, target: EventTarget | null, headerName: any) {
        if (!currentOrder || currentOrder == "none" || currentOrder == "desc") {
            // @ts-ignore
            $(target).find(".uci-order-icon").addClass("wj-glyph-up");
            // @ts-ignore
            target.dataset.order = "asc";
            this.replaceFetchOrder(headerName, false);
        }
        else if (currentOrder == "asc") {
            // @ts-ignore
            $(target).find(".uci-order-icon").addClass("wj-glyph-down");
            // @ts-ignore
            target.dataset.order = "desc";
            this.replaceFetchOrder(headerName, true);
        }
    }

    private link_Click(evt: Event): void {
        var currentItem = evt.currentTarget;

        // @ts-ignore
        var data = currentItem.dataset;
        var recordId = data.recordId;
        var recordLogicalName = data.recordLogicalName;

        var entityFormOptions: { entityName?: string, entityId?: string, openInNewWindow?: boolean } = {};
        entityFormOptions.entityName = recordLogicalName;
        entityFormOptions.entityId = recordId;
        entityFormOptions.openInNewWindow = window.event.ctrlKey;

        // Open the form.
        Xrm.Navigation.openForm(entityFormOptions);
    }

    private async loadPageFirst(): Promise<any> {
        this._headers = new Array<string>();
        // @ts-ignore
        var fetchXML = this._context.parameters.FetchXML.attributes.LogicalName;

        this._fetchXML = Xrm.Page.getAttribute(fetchXML).getValue();

        if (this._fetchXML) {
            await this.LoadPage();
        }
    }

    private async LoadPage() {
        this._fetchXML = this._fetchXML.replace(/"/g, "'");
        this.getPrimaryEntityName();

        var entityDefinition = await this.getPluralName(this._primaryEntityName);

        await this.SetVarsLoadPage(entityDefinition);

        let entitiesInFetch: Array<{
            EntityLogicalName: string;
            EntityDisplayName: string;
            Alias: string;
        }> = new Array();

        entitiesInFetch = this.getLinkedEntitiesFromFetch(this._jsonFetch.fetch.entity["link-entity"], entitiesInFetch);
        entitiesInFetch.push({ EntityLogicalName: this._jsonFetch.fetch.entity._attributes.name, EntityDisplayName: "", Alias: "" });

        await this.GetHeaderNames(entitiesInFetch);

        this._entitiesInFetch = entitiesInFetch;
        this.formatFetchXML();

        if (this._primaryEntityNamePlural && this._fetchXML && this._headerDisplayNames) {
            this.loadTable();
            this.loadPage();
        }
    }

    private async SetVarsLoadPage(entityDefinition: { EntitySetName: string; PrimaryNameAttribute: string; }) {
        this._primaryEntityNamePlural = entityDefinition.EntitySetName;
        this._primaryFieldName = entityDefinition.PrimaryNameAttribute;
        this._jsonFetch = JSON.parse(convert.xml2json(this._fetchXML, { compact: true, spaces: 2 }));
        this._totalNumberOfRecords = await this.getTotalNumberOfRecords(this._fetchXML);
    }

    private async GetHeaderNames(entitiesInFetch: { EntityLogicalName: string; EntityDisplayName: string; Alias: string; }[]) {
        for (var i in entitiesInFetch) {
            var itemDefinition = await this.getPluralName(entitiesInFetch[i].EntityLogicalName);
            // @ts-ignore
            entitiesInFetch[i].EntityDisplayName = itemDefinition.DisplayName.UserLocalizedLabel.Label;
            let headers = await this.getAttributeDisplayNames(entitiesInFetch[i].EntityLogicalName, entitiesInFetch[i].Alias);
            this.PushHeaderNames(headers, entitiesInFetch, i);
        }
    }

    private PushHeaderNames(headers: { LogicalName: string; DisplayName: string; }[], entitiesInFetch: { EntityLogicalName: string; EntityDisplayName: string; Alias: string; }[], i: string) {
        for (var j = 0; j < headers.length; j++) {
            let suffix = "";
            if (entitiesInFetch[i].Alias) {
                suffix = " (" + entitiesInFetch[i].EntityDisplayName + ")";
                headers[j].DisplayName = headers[j].DisplayName + suffix;
            }
            this._headerDisplayNames.push(headers[j]);
        }
    }

    private getLinkedEntitiesFromFetch(linkEntity: any, ret: Array<{ EntityLogicalName: string, EntityDisplayName: string, Alias: string }>): Array<{ EntityLogicalName: string; EntityDisplayName: string; Alias: string; }> {
        if (linkEntity) {
            if (linkEntity.length && linkEntity.length > 0) {
                for (var i = 0; i < linkEntity.length; i++) {
                    this.getLinkedEntitiesFromFetch(linkEntity[i], ret);
                }
            } else {
                let entityLogical = linkEntity._attributes.name;
                let alias = linkEntity._attributes.alias;

                ret.push({ EntityLogicalName: entityLogical, Alias: alias, EntityDisplayName: "" });

                this.getLinkedEntitiesFromFetch(linkEntity["link-entity"], ret);
            }
        }

        return ret;
    }

    private getPrimaryEntityName(): void {
        // @ts-ignore
        var filter = this._fetchXML.matchAll(/<entity name='(.*?)'>/g).next();

        if (filter && filter.value && filter.value[1]) {
            this._primaryEntityName = filter.value[1];
        }
    }

    private getTotalNumberOfRecords(fetchXML: string): Promise<number> {
        var fetch = this.getCountFetch();

        return new Promise(resolve => {
            this._webApi.executeFetchXml(this._primaryEntityNamePlural, fetch, undefined, undefined, undefined, undefined)
                .then((data) => {
                    resolve(data.value[0].recordcount);
                }).catch((e) => {
                    debugger;
                });
        });
    }

    private replaceFetchOrder(attrName: string, isDesc: boolean): void {
        this._jsonFetch.fetch._attributes.count = 5;
        this._jsonFetch.fetch.entity.order._attributes.attribute = attrName;
        this._jsonFetch.fetch.entity.order._attributes.descending = isDesc ? "true" : "false";

        var options = { ignoreComment: true, compact: true };

        this._fetchXML = convert.js2xml(this._jsonFetch, options);
        this.loadPage();
    }

    private getCountFetch(): string {
        var json = JSON.parse(convert.xml2json(this._fetchXML, { compact: true, spaces: 2 }));
        json.fetch.entity._attributes.aggregate = true;

        this.removeCountFetchAttributes(json);

        json.fetch.entity.attribute.push({ _attributes: { name: "createdon", alias: "recordcount", aggregate: "count" } });

        this.removeCountFetchLinkAttributes(json.fetch.entity["link-entity"]);

        json.fetch._attributes.aggregate = true;
        delete json.fetch.entity.order;

        var options = { ignoreComment: true, compact: true };
        return convert.js2xml(json, options);
    }

    private removeCountFetchAttributes(json: any) {
        if (json.fetch.entity.attribute && json.fetch.entity.attribute.length) {
            for (var i = 0; i < json.fetch.entity.attribute.length; i++) {
                if (json.fetch.entity.attribute[i]) {
                    delete json.fetch.entity.attribute[i];
                }
            }
            json.fetch.entity.attribute.length = 0;
        }
        else {
            if (json.fetch.entity.attribute) {
                delete json.fetch.entity.attribute;
            }
        }
    }

    private removeCountFetchLinkAttributes(linkEntity: any): void {
        // @ts-ignore
        if (linkEntity) {
            if (linkEntity.length && linkEntity.length > 0) {
                for (var j = 0; j < linkEntity.length; j++) {
                    this.removeCountFetchLinkAttributes(linkEntity[j]);
                }
            } else {
                if (linkEntity.attribute && linkEntity.attribute.length) {
                    for (var j = 0; j < linkEntity.attribute.length; j++) {
                        if (linkEntity.attribute[j]) {
                            delete linkEntity.attribute[j];
                        }
                    }

                    linkEntity.attribute.length = 0;
                } else {
                    if (linkEntity.attribute) {
                        delete linkEntity.attribute;
                    }
                }

                this.removeCountFetchLinkAttributes(linkEntity["link-entity"]);
            }
        }
    }

    private formatFetchXML(): void {
        var regex = /count='[^"]'/g;
        this._fetchXML = this._fetchXML.replace(regex, '');
        regex = /page='[^"]'/g;
        this._fetchXML = this._fetchXML.replace(regex, '');

        this._fetchXML = this._fetchXML.replace("fetch", "fetch count='5'");
    }

    private loadPage(): void {
        var daddy = this;

        this._webApi.executeFetchXml(this._primaryEntityNamePlural, this._fetchXML, "*", this._currentPageNumber, undefined, undefined)
            .then((data) => {
                daddy.loadGrid(data);
            }).catch((e) => {
                debugger;
            });
    }

    private loadTable(): void {
        this._grid = document.createElement("table");
        this._grid.setAttribute("id", "tbl_records");
        this._grid.className = "uci-table";

        var tableHeader = this.getHeader();

        var tbody = document.createElement("tbody");
        var tfoot = document.createElement("tfoot");

        this._grid.appendChild(tableHeader);
        this._grid.appendChild(tbody);
        this._grid.appendChild(tfoot);
    }

    private selectGridOrder(): void {
        var attr = this._jsonFetch.fetch.entity.order._attributes.attribute;
        var desc = this._jsonFetch.fetch.entity.order._attributes.descending;
        var isDesc = desc == "true" ? true : false;

        $(".uci-thead-th[data-logical-name='" + attr + "'").data("order", desc);

        if (isDesc) {
            $(".uci-thead-th[data-logical-name='" + attr + "'").find(".uci-order-icon").addClass("wj-glyph-down");
        } else {
            $(".uci-thead-th[data-logical-name='" + attr + "'").find(".uci-order-icon").addClass("wj-glyph-up");
        }
    }

    private loadGrid(data: any): void {
        var tBody = this.getBody(data);
        var tFoot = this.getFooter(data);

        this._grid.children[1].remove();
        this._grid.children[1].remove();

        this._grid.appendChild(tBody);
        this._grid.appendChild(tFoot);

        if (this._container.hasChildNodes() && this._container.firstChild != null) {
            this._container.firstChild.remove();
        }

        this._container.appendChild(this._grid);
        this.selectGridOrder();
        // @ts-ignore
        //$(".uci-table").colResizable();
    }

    private getFooter(data: any): HTMLElement {
        let footer = document.createElement("tfoot");
        let footerTR = document.createElement("tr");
        let footerTDCount = document.createElement("td");
        let footerTDPaging = document.createElement("td");

        footer.className = "uci-footer";
        footerTR.className = "uci-footer-tr";
        footerTDCount.className = "uci-footer-td uci-footer-td-left";
        footerTDPaging.className = "uci-footer-td uci-footer-td-right";

        footerTDPaging.colSpan = this.getPagingColSpan();

        let countText = this.getCountText(data);
        footerTDCount.textContent = countText;

        let footerPagingDiv = this.getPagingDiv(data);
        footerTDPaging.appendChild(footerPagingDiv);

        footerTR.appendChild(footerTDCount);
        footerTR.appendChild(footerTDPaging);
        footer.appendChild(footerTR);

        return footer;
    }

    private getPagingColSpan(): number {
        var headerCount = this._headers.length;
        return headerCount - 1;
    }

    private getPagingDiv(data: any): HTMLElement {
        let pagingDiv = document.createElement("div");
        pagingDiv.className = "uci-footer-paging";

        let firstButton = document.createElement("button");
        let previousButton = document.createElement("button");
        let nextButton = document.createElement("button");

        this.setPagingDivEvents(firstButton, previousButton, nextButton);

        let pageSpan = this.setPagingDivSpan();

        this.setPagingDivButtonClasses(firstButton, previousButton, nextButton);

        this.checkPagingDivFirstPage(firstButton, previousButton);

        let finalItem = this._currentPageNumber * this._itemsPerPage;
        if (finalItem >= this._totalNumberOfRecords) {
            nextButton.disabled = true;
        }

        this.appendPagingDivChildren(pagingDiv, firstButton, previousButton, pageSpan, nextButton);

        return pagingDiv;
    }

    private appendPagingDivChildren(pagingDiv: HTMLDivElement, firstButton: HTMLButtonElement, previousButton: HTMLButtonElement, pageSpan: HTMLSpanElement, nextButton: HTMLButtonElement) {
        pagingDiv.appendChild(firstButton);
        pagingDiv.appendChild(previousButton);
        pagingDiv.appendChild(pageSpan);
        pagingDiv.appendChild(nextButton);
    }

    private checkPagingDivFirstPage(firstButton: HTMLButtonElement, previousButton: HTMLButtonElement) {
        if (this._currentPageNumber == 1) {
            firstButton.disabled = true;
            previousButton.disabled = true;
        }
    }

    private setPagingDivSpan() {
        let pageSpan = document.createElement("span");
        pageSpan.textContent = "Page " + this._currentPageNumber;
        pageSpan.className = "uci-span-page";
        return pageSpan;
    }

    private setPagingDivButtonClasses(firstButton: HTMLButtonElement, previousButton: HTMLButtonElement, nextButton: HTMLButtonElement) {
        firstButton.className = "symbolFont FirstPageButton-symbol uci-first-button";
        previousButton.className = "symbolFont BackButton-symbol uci-previous-button";
        nextButton.className = "symbolFont Forward-symbol uci-next-button";
    }

    private setPagingDivEvents(firstButton: HTMLButtonElement, previousButton: HTMLButtonElement, nextButton: HTMLButtonElement) {
        firstButton.addEventListener("click", this._firstButton_Click);
        previousButton.addEventListener("click", this._previousButton_Click);
        nextButton.addEventListener("click", this._nextButton_Click);
    }

    private getCountText(data: any): string {
        let initialItem = (this._currentPageNumber - 1) * this._itemsPerPage + 1;
        let finalItem = this._currentPageNumber * this._itemsPerPage;
        finalItem = finalItem > this._totalNumberOfRecords ? this._totalNumberOfRecords : finalItem;

        return initialItem + " - " + finalItem + " of " + this._totalNumberOfRecords;
    }

    private getBody(data: any): HTMLElement {
        let body = document.createElement("tbody");
        body.className = "uci-tbody";

        for (let i = 0; i < this._itemsPerPage; i++) {
            if (data.value[i]) {
                this.getBodyTR(data, i, body);
            } else {
                break;
            }
        }

        return body;
    }

    private getBodyTR(data: any, i: number, body: HTMLTableSectionElement) {
        let recordId = "";
        let bodyTR = document.createElement("tr");
        bodyTR.className = "uci-tbody-tr no-text-select";

        if (data.value[i][this._primaryEntityName + "id"]) {
            recordId = data.value[i][this._primaryEntityName + "id"];
        }

        bodyTR.dataset.recordId = recordId;
        bodyTR.dataset.recordLogicalName = this._primaryEntityName;
        bodyTR.addEventListener("dblclick", this._link_Click);

        this.getBodyTD(data, i, recordId, bodyTR);

        body.appendChild(bodyTR);
    }

    private getBodyTD(data: any, i: number, recordId: string, bodyTR: HTMLTableRowElement) {
        for (var j in this._headers) {
            let currentHeader = this._headers[j];
            let bodyTD = document.createElement("td");
            bodyTD.className = "uci-tbody-td";

            let lookupName = "_" + currentHeader + "_value";
            let lookupValue = data.value[i][lookupName];

            var formattedValue = data.value[i][currentHeader + "@OData.Community.Display.V1.FormattedValue"];
            var textContent: string = formattedValue ? formattedValue : data.value[i][currentHeader];

            textContent = this.getBodyTextContent(currentHeader, textContent, recordId, bodyTD, lookupValue, data, i, lookupName);

            bodyTR.appendChild(bodyTD);
        }
    }

    private getBodyTextContent(currentHeader: string, textContent: string, recordId: string, bodyTD: HTMLTableDataCellElement, lookupValue: any, data: any, i: number, lookupName: string) {
        if (currentHeader == this._primaryFieldName) {
            this.getPrimaryFieldTD(textContent, recordId, bodyTD);
        }
        else if (lookupValue) {
            this.getLookupFieldTD(data.value[i], lookupName, bodyTD);
        }
        else {
            textContent = textContent ? textContent : "---";
            var finalTextContent = textContent.length >= 30 ? textContent.substr(0, 30) + "..." : textContent;
            bodyTD.textContent = finalTextContent;
            bodyTD.title = textContent;
        }
        return textContent;
    }

    private getLookupFieldTD(dataItem: any, lookupName: string, bodyTD: HTMLTableDataCellElement): void {
        let lookupId = dataItem[lookupName];
        let lookupLogicalName = dataItem[lookupName + "@Microsoft.Dynamics.CRM.lookuplogicalname"];
        let lookupDisplayName = dataItem[lookupName + "@OData.Community.Display.V1.FormattedValue"];

        let aElement = document.createElement("a");
        aElement.href = "#";
        aElement.textContent = lookupDisplayName;
        aElement.dataset.recordId = lookupId;
        aElement.dataset.recordLogicalName = lookupLogicalName;
        aElement.addEventListener("click", this._link_Click);

        bodyTD.appendChild(aElement);
    }

    private getPrimaryFieldTD(textContent: string, recordId: string, bodyTD: HTMLTableDataCellElement) {
        let aElement = document.createElement("a");
        aElement.href = "#";
        aElement.textContent = textContent ? textContent : "---";
        aElement.dataset.recordId = recordId;
        aElement.dataset.recordLogicalName = this._primaryEntityName;
        aElement.addEventListener("click", this._link_Click);
        bodyTD.appendChild(aElement);
    }

    private getHeader(): HTMLElement {
        let header = document.createElement("thead");
        let headerTR = document.createElement("tr");

        header.className = "uci-thead";
        headerTR.className = "uci-thead-tr no-text-select";

        let attributes: Array<string> = new Array();
        attributes = this.getAttributesFromFetch(attributes);
        attributes = this.getLinkedAttributesFromFetch(this._jsonFetch.fetch.entity["link-entity"], attributes);

        for (var item in attributes) {
            let headerTH = document.createElement("th");
            let headerTHDiv = document.createElement("div");
            let headerLogicalName = attributes[item];

            headerTHDiv.dataset.logicalName = headerLogicalName;
            headerTHDiv.dataset.order = "none";

            headerTHDiv.addEventListener("click", this._th_Click);

            if (headerLogicalName == this._primaryEntityName + "id") {
                continue;
            }

            // @ts-ignore
            let headerDisplayNameItem = this._headerDisplayNames.find(item => item.LogicalName == headerLogicalName);

            if (headerDisplayNameItem) {
                let headerDisplayName = headerDisplayNameItem.DisplayName;

                headerTHDiv.className = "uci-thead-th";

                headerTHDiv.textContent = headerDisplayName;
                headerTHDiv.title = headerDisplayName;
                this._headers.push(attributes[item]);

                let orderIcon = document.createElement("div");
                orderIcon.className = "uci-order-icon";

                headerTHDiv.appendChild(orderIcon);

                headerTH.appendChild(headerTHDiv);
                headerTR.appendChild(headerTH);
            }
        }

        header.appendChild(headerTR);

        return header;
    }

    private getAttributesFromFetch(ret: string[]): string[] {
        let attributes = this._jsonFetch.fetch.entity.attribute;

        for (var i = 0; i < attributes.length; i++) {
            ret.push(attributes[i]._attributes.name);
        }

        return ret;
    }

    private getLinkedAttributesFromFetch(linkEntity: any, ret: Array<string>): any {
        // @ts-ignore
        if (linkEntity) {
            if (linkEntity.length && linkEntity.length > 0) {
                for (var i = 0; i < linkEntity.length; i++) {
                    this.getLinkedAttributesFromFetch(linkEntity[i], ret);
                }
            } else {
                var alias = linkEntity._attributes.alias + ".";

                // @ts-ignore
                //var displayNameItem = this._entitiesInFetch.find(x => x.EntityLogicalName == name);
                //var displayName = displayNameItem ? " (" + displayNameItem.EntityDisplayName + ")" : "";

                var displayName = "";

                if (linkEntity.attribute && linkEntity.attribute.length) {
                    for (var i = 0; i < linkEntity.attribute.length; i++) {
                        if (linkEntity.attribute[i]) {
                            ret.push(alias + linkEntity.attribute[i]._attributes.name + displayName);
                        }
                    }
                } else {
                    if (linkEntity.attribute) {
                        ret.push(alias + linkEntity.attribute._attributes.name + displayName);
                    }
                }

                this.getLinkedAttributesFromFetch(linkEntity["link-entity"], ret);
            }
        }

        return ret;
    }

    private getPluralName(logicalName: string): Promise<{ EntitySetName: string, PrimaryNameAttribute: string }> {
        var request = {
            collection: "EntityDefinitions",
            select: ["EntitySetName", "LogicalName", "PrimaryNameAttribute", "DisplayName"],
            filter: "LogicalName eq '" + logicalName + "'"
        };

        return new Promise(resolve => {
            this._webApi.retrieveMultipleRequest(request).then(function (data: any) {
                var ret: { EntitySetName: string, PrimaryNameAttribute: string, DisplayName: string };
                ret = { EntitySetName: data.value[0].EntitySetName, PrimaryNameAttribute: data.value[0].PrimaryNameAttribute, DisplayName: data.value[0].DisplayName };

                resolve(ret);
            }).catch(function (error: any) {
                debugger;
            });
        });
    }

    private getAttributeDisplayNames(primaryEntityName: string, alias: string): Promise<Array<{ LogicalName: string, DisplayName: string }>> {
        var request = {
            collection: 'EntityDefinitions',
            key: "LogicalName='" + primaryEntityName + "'",
            navigationProperty: 'Attributes',
            select: ['LogicalName', 'SchemaName', 'DisplayName']
        };

        return new Promise(resolve => {
            this._webApi.retrieveRequest(request).then(function (data: any) {
                let ret = new Array();

                if (data && data.value) {
                    ret = $.map(data.value, (item) => {
                        let logicalName = alias ? alias + "." + item.LogicalName : item.LogicalName;

                        if (item && item.DisplayName && item.DisplayName.UserLocalizedLabel) {
                            return { LogicalName: logicalName, DisplayName: item.DisplayName.UserLocalizedLabel.Label };
                        } else {
                            return { LogicalName: logicalName, DisplayName: item.LogicalName };
                        }
                    });
                }

                resolve(ret);
            }).catch(function (error: any) {
                debugger;
            });
        });
    }

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        if (!this._lock) {
            this._lock = true;
            this.loadPageFirst();

            // TODO: Quick fix
            var daddy = this;
            setTimeout(() => { daddy._lock = false; }, 1000);
        }
    }

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
    public getOutputs(): IOutputs {
        return {};
    }

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}