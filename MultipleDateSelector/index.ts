import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class MultipleDateSelector implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _selectedDates: Date[] = [];
    private _combinedValue = "";
    private _lastInputValue = "";
    private _isRemoving = false;
    private _dateFormat: string = "0"; // 0 = FullDate (browser locale), 1 = YearOnly

    // UI Elements
    private _dateInput: HTMLInputElement;
    private _selectedDatesContainer: HTMLDivElement;

    // Constructor intentionally empty - initialization happens in init method

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;

        // Load date format configuration
        const dateFormatValue = context.parameters.dateFormat.raw;
        this._dateFormat = dateFormatValue !== null && dateFormatValue !== undefined ? dateFormatValue.toString() : "0";

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
                if (date && !isNaN(date.getTime()) && input.validity.valid) {
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

        this.updateSelectedDatesDisplay();
    }

    private addDate(): void {
        if (this._dateInput.value && !this._isRemoving) {
            const selectedDate = this.parseISODateAsLocal(this._dateInput.value);

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
            // Format display based on selected format
            if (this._dateFormat === "1") {
                // Year only
                dateText.textContent = date.getFullYear().toString();
            } else {
                dateText.textContent = date.toLocaleDateString();
            }
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
        if (this._dateFormat === "1") {
            // Year only format
            const years = this._selectedDates
                .map(date => date.getFullYear().toString())
                .filter((value, index, self) => self.indexOf(value) === index)
                .sort();
            this._combinedValue = JSON.stringify(years);
        } else {
            this._combinedValue = JSON.stringify(this._selectedDates.map(date => date.toLocaleDateString()));
        }

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
        // Update date format if it changed
        const dateFormatValue = context.parameters.dateFormat.raw;
        const newFormat = dateFormatValue !== null && dateFormatValue !== undefined ? dateFormatValue.toString() : "0";
        if (newFormat !== this._dateFormat) {
            this._dateFormat = newFormat;
            this.updateSelectedDatesDisplay();
            this.updateCombinedValue();
        }

        // Only update if the external value is different from our current combined value
        const externalValue = context.parameters.value.raw || "";
        if (externalValue !== this._combinedValue) {
            if (this.parseExternalValue(externalValue)) {
                this.updateCombinedValue();
            }
        }
    }

    private parseExternalValue(value: string): boolean {
        let parts: string[];
        const trimmedValue = value.trim();

        if (trimmedValue.startsWith("[")) {
            let parsedValue: unknown;
            try {
                parsedValue = JSON.parse(trimmedValue);
            } catch (error) {
                console.error("Unable to parse saved date array:", error);
                return false;
            }

            if (!Array.isArray(parsedValue) || !parsedValue.every(part => typeof part === "string")) {
                console.error("Saved date value must be a JSON array of strings.");
                return false;
            }

            parts = parsedValue;
        } else {
            // Continue accepting values saved by older versions of the control.
            parts = trimmedValue ? trimmedValue.split(",").map(part => part.trim()) : [];
        }

        this._selectedDates = [];
        parts.forEach(part => {
            if (part) {
                const date = this.parseDate(part);
                if (date && !isNaN(date.getTime())) {
                    const dateExists = this._selectedDates.some(existingDate =>
                        existingDate.toDateString() === date.toDateString()
                    );
                    if (!dateExists) {
                        this._selectedDates.push(date);
                    }
                }
            }
        });

        this._selectedDates.sort((a, b) => a.getTime() - b.getTime());
        this.updateSelectedDatesDisplay();
        return true;
    }

    private parseDate(dateString: string): Date | null {
        // Check if it's year-only format (4 digits)
        if (/^\d{4}$/.test(dateString)) {
            const year = parseInt(dateString, 10);
            if (year >= 1900 && year <= 2100) {
                return new Date(year, 0, 1); // January 1st of the year in local time
            }
            return null;
        }

        // Check if it's ISO format (YYYY-MM-DD)
        if (/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
            return this.parseISODateAsLocal(dateString);
        }

        const localizedDate = this.parseBrowserLocaleDate(dateString);
        if (localizedDate) {
            return localizedDate;
        }

        // Try multiple date formats for better compatibility
        const slashMatch = dateString.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (slashMatch) {
            const first = parseInt(slashMatch[1], 10);
            const second = parseInt(slashMatch[2], 10);
            const year = parseInt(slashMatch[3], 10);

            // If first number > 12, it must be day (European format DD/MM/YYYY)
            if (first > 12 && second <= 12) {
                const date = this.parseISODateAsLocal(`${year}-${slashMatch[2]}-${slashMatch[1]}`);
                if (date && !isNaN(date.getTime()) && this.isValidDate(date, first, second, year)) {
                    return date;
                }
            }

            // If second number > 12, it must be day (US format interpretation would be invalid)
            if (second > 12 && first <= 12) {
                const date = this.parseISODateAsLocal(`${year}-${slashMatch[1]}-${slashMatch[2]}`);
                if (date && !isNaN(date.getTime()) && this.isValidDate(date, second, first, year)) {
                    return date;
                }
            }

            // Both could be valid (both <= 12), try US format first, then European
            const usFormat = this.parseISODateAsLocal(`${year}-${slashMatch[1]}-${slashMatch[2]}`);
            if (usFormat && !isNaN(usFormat.getTime()) && this.isValidDate(usFormat, second, first, year)) {
                return usFormat;
            }

            const euFormat = this.parseISODateAsLocal(`${year}-${slashMatch[2]}-${slashMatch[1]}`);
            if (euFormat && !isNaN(euFormat.getTime()) && this.isValidDate(euFormat, first, second, year)) {
                return euFormat;
            }
        }

        return null;
    }

    private parseBrowserLocaleDate(dateString: string): Date | null {
        const formatParts = new Intl.DateTimeFormat().formatToParts(new Date(2006, 10, 22));
        const fieldOrder = formatParts
            .filter(part => part.type === "day" || part.type === "month" || part.type === "year")
            .map(part => part.type);

        if (fieldOrder.length !== 3) {
            return null;
        }

        const localeDigits = Array.from(new Intl.NumberFormat(undefined, { useGrouping: false }).format(9876543210));
        const digitMap = new Map(localeDigits.map((digit, index) => [digit, String(9 - index)]));
        const normalizedValue = Array.from(dateString, character => digitMap.get(character) ?? character).join("");
        const values = normalizedValue.match(/\d+/g);

        if (!values || values.length !== fieldOrder.length) {
            return null;
        }

        const dateParts: Partial<Record<"day" | "month" | "year", number>> = {};
        fieldOrder.forEach((field, index) => {
            if (field === "day" || field === "month" || field === "year") {
                dateParts[field] = Number(values[index]);
            }
        });

        const { day, month, year } = dateParts;
        if (day === undefined || month === undefined || year === undefined) {
            return null;
        }

        const date = this.parseISODateAsLocal(
            `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`
        );

        return this.isValidDate(date, day, month, year) ? date : null;
    }

    private isValidDate(date: Date, day: number, month: number, year: number): boolean {
        // Verify the parsed date matches the input values
        return date.getDate() === day && 
               date.getMonth() === month - 1 && 
               date.getFullYear() === year;
    }

    private parseISODateAsLocal(dateString: string): Date {
        // Parse YYYY-MM-DD format as local time, not UTC
        const [year, month, day] = dateString.split('-').map(Number);
        return new Date(year, month - 1, day);
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