import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { IDropdownOption } from "office-ui-fabric-react";
import DurationControl from "./DurationControl";
import * as React from "react";
import * as ReactDOM from "react-dom";

enum Duration_Type {
	Minute,
	Hour,
	Day
}

export class CustomDurationPCF implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private container: HTMLDivElement;
	private notifyOutputChanged: () => void;
	private currentValue: number | null;
	private availableValues: number[];

	constructor() {

	}

	private getDurationTextByPart(value: number, part: Duration_Type): string {
		switch (part) {
			case Duration_Type.Day:
				const days = value / 1440;
				return days.toString() + " " + (1 != days ? "days" : "day");
			case Duration_Type.Hour:
				const hours = value / 60;
				return hours.toString() + " " + (1 != hours ? "hours" : "hour");
			case Duration_Type.Minute:
				return value + " " + (1 != value ? "minutes" : "minute");
		}
		return "";
	}

	private getDurationText(value: number): string {
		var minutes = Math.round(value)
			, hour = 60
			, day = 24 * hour
			, days = minutes / day
			, hours = minutes / hour;
		return minutes >= day && days === Math.round(100 * days) / 100 ? this.getDurationTextByPart(minutes, Duration_Type.Day) : minutes >= hour && hours === Math.round(100 * hours) / 100 ? this.getDurationTextByPart(minutes, Duration_Type.Hour) : this.getDurationTextByPart(minutes, Duration_Type.Minute);
	}

	public init(context: ComponentFramework.Context<IInputs>,
		notifyOutputChanged: () => void,
		state: ComponentFramework.Dictionary,
		container: HTMLDivElement) {
		this.container = container;
		this.notifyOutputChanged = notifyOutputChanged;

		let availableValuesString = context.parameters.availableValues.raw;

		if (availableValuesString == null) {
			container.innerHTML = "Property 'Available Values' is blank, configure it please for correct work";
			return;
		}

		this.availableValues = availableValuesString.split("|").map(t => parseInt(t));

		this.renderControl(context);
	}

	public updateView(context: ComponentFramework.Context<IInputs>): void {
		this.renderControl(context);
	}

	private renderControl(context: ComponentFramework.Context<IInputs>) {
		let availableValue = this.availableValues.slice();

		let _currentValue = context.parameters.value.raw;

		if (_currentValue != null && !isNaN(_currentValue) && !availableValue.some(t => t === _currentValue)) {
			availableValue.push(_currentValue);
			availableValue.sort((a, b) => a - b);
		}

		const options = availableValue.map(t => ({
			key: t,
			text: this.getDurationText(t)
		}));

		const dropDownProperties = {
			options: options,
			selectedKey: _currentValue,
			onSelectedChanged: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
				this.currentValue = option == undefined ? null : <number>option.key;
				this.notifyOutputChanged();
			}
		};

		ReactDOM.render(React.createElement(DurationControl, dropDownProperties), this.container);
	}

	public getOutputs(): IOutputs {
		return {
			value: this.currentValue == null ? undefined : this.currentValue
		};
	}

	public destroy(): void {
		ReactDOM.unmountComponentAtNode(this.container);
	}
}