import powerbi from "powerbi-visuals-api";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;

export class Visual implements IVisual {
    private readonly host: powerbi.extensibility.visual.IVisualHost;
    private readonly table: HTMLTableElement;
    private readonly selectionManager: powerbi.extensibility.ISelectionManager;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = this.host.createSelectionManager();

        if (document) {
            this.table = document.createElement("table");
            options.element.appendChild(this.table);
        }
    }

    public update(options: VisualUpdateOptions) {
        const dataview = options.dataViews.at(0);

        if (dataview == null || dataview.table == null) {
            return;
        }

        const rowNodes = [];

        const headerRow = document.createElement("tr");
        for (const column of dataview.table.columns) {
            const headerCell = document.createElement("th");
            headerCell.textContent = column.displayName;
            headerRow.appendChild(headerCell);
        }

        rowNodes.push(headerRow);

        for (let rowIndex = 0; rowIndex < dataview.table.rows.length; rowIndex++) {
            const row = dataview.table.rows[rowIndex];
            const rowElement = document.createElement("tr");

            for (const item of row) {
                const valueCell = document.createElement("td");
                valueCell.textContent = String(item);
                rowElement.appendChild(valueCell);
            }

            const selectionId = this.host.createSelectionIdBuilder()
                .withTable(dataview.table, rowIndex)
                .createSelectionId();

            rowElement.addEventListener("contextmenu", (event) => {
                this.selectionManager.showContextMenu(
                    selectionId,
                    { x: event.x, y: event.y },
                    "values" // Not sure what this should be, does not seem to make a difference
                );

                // Prevent the browser context-menu from showing
                event.preventDefault();
                return false;
            })

            // Add event listener on hovering to show or hide tooltips
            rowElement.addEventListener("mouseenter", (event) => {
                if (this.host.tooltipService.enabled()) {

                    const tooltipInfo: powerbi.extensibility.VisualTooltipDataItem[] = row.map((value, index) => {
                        return {
                            displayName: dataview.table.columns[index].displayName,
                            value: String(value)
                        }
                    });

                    this.host.tooltipService.show({
                        coordinates: [event.x, event.y],
                        isTouchEvent: true,
                        dataItems: tooltipInfo,
                        identities: [selectionId]
                    });
                }
            });

            rowNodes.push(rowElement);
        }

        this.table.replaceChildren(...rowNodes)
    }
}