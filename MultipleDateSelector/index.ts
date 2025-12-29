import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class MultipleDateSelector implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _selectedDates: Date[] = [];
    private _combinedValue = "";
    private _lastInputValue = "";
    private _isRemoving = false;

    // UI Elements
    private _dateInput: HTMLInputElement;
    private _selectedDatesContainer: HTMLDivElement;

    // Constructor intentionally empty - initialization happens in init method

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;

        this.createUI();
        this.loadExistingValues();
        this.updateCombinedValue();

    }

    private createUI(): void {
        // Main container
        
        this._container.style.fontFamily = "Segoe UI, sans-serif";

        // Title


        // Date input section
        const dateSection = document.createElement("div");
        dateSection.style.marginBottom = "15px";

        const dateInputContainer = document.createElement("div");
        dateInputContainer.style.display = "flex";
        dateInputContainer.style.gap = "10px";
        dateInputContainer.style.alignItems = "center";

        this._dateInput = document.createElement("input");
        this._dateInput.type = "date";
        this._dateInput.style.padding = "8px";
        this._dateInput.style.border = "none";
        this._dateInput.style.borderRadius = "4px";
        this._dateInput.style.fontSize = "14px";
        this._dateInput.style.width = "100%";
        this._dateInput.style.backgroundColor = "rgba(249, 249, 249)";
        this._dateInput.min = "1900-01-01";

        this._dateInput.addEventListener('change', (e) => {
            // Only add if we're not in the middle of a remove operation
             const input = e.target as HTMLInputElement; // ✅ cast to correct type
            if (!this._isRemoving && this._dateInput.value && this._dateInput.value !== this._lastInputValue) {
                 const date = this.parseDate(this._dateInput.value);
                    if (date && !isNaN(date.getTime())&& input.validity.valid){
 this.addDate();
                this._lastInputValue = this._dateInput.value;
                    }
               
            }
        });

        dateInputContainer.appendChild(this._dateInput);

        dateSection.appendChild(dateInputContainer);
        this._container.appendChild(dateSection);

        // Selected dates section
        const selectedSection = document.createElement("div");
        selectedSection.style.marginBottom = "15px";

        const selectedLabel = document.createElement("label");
        selectedLabel.textContent = "Selected Dates:";
        selectedLabel.style.display = "block";
        selectedLabel.style.marginBottom = "5px";
        selectedLabel.style.fontWeight = "600";
        selectedSection.appendChild(selectedLabel);

        this._selectedDatesContainer = document.createElement("div");
        this._selectedDatesContainer.style.minHeight = "40px";

        this._selectedDatesContainer.style.borderRadius = "4px";
        this._selectedDatesContainer.style.padding = "8px";
        this._selectedDatesContainer.style.backgroundColor = "#f9f9f9";
        selectedSection.appendChild(this._selectedDatesContainer);

        this._container.appendChild(selectedSection);


        // const addButton = document.createElement("button");
        // addButton.textContent = "Add";
        // addButton.style.padding = "8px 12px";
        // addButton.style.backgroundColor = "#0078d4";
        // addButton.style.color = "white";
        // addButton.style.border = "none";
        // addButton.style.borderRadius = "4px";
        // addButton.style.cursor = "pointer";
        // addButton.style.width = "20%";

        // addButton.addEventListener('click', () => {
        //     if (!this._isRemoving && this._dateInput.value) {
        //         this.addDate();
        //         this._dateInput.value = "";
        //         this._lastInputValue = "";
        //     }
        // });
        // dateInputContainer.appendChild(addButton);
        this.updateSelectedDatesDisplay();
    }

    private addDate(): void {
        if (this._dateInput.value && !this._isRemoving) {
            const selectedDate = new Date(this._dateInput.value);

            // Check if date is already selected
            const dateExists = this._selectedDates.some(date =>
                date.toDateString() === selectedDate.toDateString()
            );

            if (!dateExists) {
                this._selectedDates.push(selectedDate);
                this._selectedDates.sort((a, b) => a.getTime() - b.getTime());
                this.updateSelectedDatesDisplay();
                this.updateCombinedValue();
            }
            // Clear the input after adding
            this._dateInput.value = "";
            this._lastInputValue = "";
        }
    }

    private removeDate(index: number): void {
        this._isRemoving = true;
        this._dateInput.value = "";
        this._lastInputValue = "";

        this._selectedDates.splice(index, 1);
        this.updateSelectedDatesDisplay();
        this.updateCombinedValue();

        // Add a slightly longer delay to ensure everything settles
        setTimeout(() => {
            this._isRemoving = false;
        }, 300);
    }

    private updateSelectedDatesDisplay(): void {
        this._selectedDatesContainer.innerHTML = "";

        if (this._selectedDates.length === 0) {
            const emptyMessage = document.createElement("span");
            emptyMessage.textContent = "No dates selected";
            emptyMessage.style.color = "#666";
            emptyMessage.style.fontStyle = "italic";
            this._selectedDatesContainer.appendChild(emptyMessage);
            return;
        }

        this._selectedDates.forEach((date, index) => {
            const dateTag = document.createElement("span");
            dateTag.style.display = "inline-block";
            dateTag.style.backgroundColor = "#0078d4";
            dateTag.style.color = "white";
            dateTag.style.padding = "4px 8px";
            dateTag.style.margin = "2px";
            dateTag.style.borderRadius = "12px";
            dateTag.style.fontSize = "12px";
            dateTag.style.cursor = "pointer";

            const dateText = document.createElement("span");
            dateText.textContent = date.toLocaleDateString();
            dateTag.appendChild(dateText);

            const removeButton = document.createElement("span");
            removeButton.textContent = " ×";
            removeButton.style.marginLeft = "5px";
            removeButton.style.fontWeight = "bold";
            removeButton.style.cursor = "pointer";
            removeButton.style.userSelect = "none";
            removeButton.style.padding = "2px";
            removeButton.addEventListener('click', (e) => {
                e.stopPropagation();
                e.preventDefault();
                e.stopImmediatePropagation();
                this.removeDate(index);
            });
            dateTag.appendChild(removeButton);

            this._selectedDatesContainer.appendChild(dateTag);
        });
    }

    private updateCombinedValue(): void {
        this._combinedValue = this._selectedDates
            .map(date => date.toISOString().split('T')[0]) // Store as YYYY-MM-DD
            .join(","); // Use comma as separator

        this._notifyOutputChanged();
    }


    private loadExistingValues(): void {
        // Load values when the control first initializes
        const existingValue = this._context.parameters.value.raw || "";

        // Debug logging (remove in production)
        console.log("Loading existing values:", existingValue);

        if (existingValue) {
            this.parseExternalValue(existingValue);
        }
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Only update if the external value is different from our current combined value
        const externalValue = context.parameters.value.raw || "";
        if (externalValue !== this._combinedValue) {
            this.parseExternalValue(this._combinedValue);
        }
    }

    private parseExternalValue(value: string): void {
        // Parse comma-separated dates from saved field value
        if (value && value.trim()) {
            const parts = value.split(',').map(part => part.trim());
            this._selectedDates = [];

            parts.forEach(part => {
                if (part) {
                    // Try different date formats to be more flexible
                    const date = this.parseDate(part);
                    if (date && !isNaN(date.getTime())) {
                        // Check if this date is already in the array (avoid duplicates)
                        const dateExists = this._selectedDates.some(existingDate =>
                            existingDate.toDateString() === date!.toDateString()
                        );
                        if (!dateExists) {
                            this._selectedDates.push(date);
                        }
                    }
                }
            });

            // Sort dates chronologically
            this._selectedDates.sort((a, b) => a.getTime() - b.getTime());
            this.updateSelectedDatesDisplay();
        } else {
            // Clear everything if no value
            this._selectedDates = [];
            this.updateSelectedDatesDisplay();

        }
    }
 
    private parseDate(dateString: string): Date | null {
        // Try multiple date formats for better compatibility
        const formats = [
            // Standard formats
            dateString,
            // US format MM/DD/YYYY
            dateString.replace(/(\d{1,2})\/(\d{1,2})\/(\d{4})/, '$3-$1-$2'),
            // European format DD/MM/YYYY  
            dateString.replace(/(\d{1,2})\/(\d{1,2})\/(\d{4})/, '$3-$2-$1'),
            // ISO format variants
            dateString.replace(/(\d{1,2})\/(\d{1,2})\/(\d{4})/, '$1/$2/$3'),
        ];

        for (const format of formats) {
            const date = new Date(format);
            if (!isNaN(date.getTime())) {
                return date;
            }
        }

        return null;
    }

    public getOutputs(): IOutputs {
        return {
            value: this._combinedValue
        };
    }

    public destroy(): void {
        // Clean up
    }
}