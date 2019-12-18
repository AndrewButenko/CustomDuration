import * as React from 'react';
import {Dropdown, IDropdownOption, initializeIcons} from 'office-ui-fabric-react';

interface IDurationControlProperties {
    options: IDropdownOption[];
    selectedKey: number | null;
    onSelectedChanged: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;    
}

initializeIcons();

export default class DurationControl extends React.Component<IDurationControlProperties, {}> {
    render() {
        return (
            <Dropdown
                placeHolder="--Select--"
                options={this.props.options}
                selectedKey={this.props.selectedKey}
                onChange={this.props.onSelectedChanged}
            />
        );
    }
}