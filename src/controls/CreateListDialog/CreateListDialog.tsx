import SPNameValidator, { Platform, ValidationType } from '@creativeacer/spnamevalidator/SPNameValidator';
import { escape } from '@microsoft/sp-lodash-subset';
import { css, DefaultButton, Dialog, DialogFooter, DialogType, Link, List, PrimaryButton, Spinner, TextField } from 'office-ui-fabric-react';
import * as React from 'react';
import { IList } from "../../interfaces/IList";
import styles from './CreateListDialog.module.scss';

export interface ICreateListDialogState {
    hideDialog: boolean;
    listTitle: string;
    loading: boolean;
    createdLists: IList[];
    errorMsg: string;
    error: boolean;
}
export interface ICreateListDialogProps {
    buttonLabel: string;
    dialogTitle: string;
    dialogText: string;
    saveAction: (listName: String) => Promise<IList>;
}

export class CreateListDialog extends React.Component<ICreateListDialogProps, ICreateListDialogState> {
    private _list: List;
    private _validator: SPNameValidator;
    constructor(props: ICreateListDialogProps, state: ICreateListDialogState) {
        super(props);

        this.state = {
            hideDialog: true,
            listTitle: "",
            loading: false,
            createdLists: [],
            errorMsg: "",
            error: false
        };

        this._validator = new SPNameValidator(Platform["SharePoint Online"]);
    }

    public render() {
        return (
            <div>
                <DefaultButton
                    secondaryText='Opens the Create List Dialog'
                    onClick={this._showDialog}
                    text={this.props.buttonLabel}
                />
                <Dialog
                    hidden={this.state.hideDialog}
                    onDismiss={this._closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: this.props.dialogTitle,
                        subText: this.state.loading ? null : this.props.dialogText
                    }}
                    modalProps={{
                        isBlocking: this.state.loading ? true : false,
                        containerClassName: 'ms-dialogMainOverride'
                    }}>

                    {!this.state.loading ? (
                        <div>
                            {this.state.createdLists.length > 0 &&
                                <div className={styles["ms-ListScrollingExample-container"]} data-is-scrollable={true}>
                                    <p className="ms-font-m">Created lists:</p>
                                    <List
                                        ref={this._resolveList}
                                        items={this.state.createdLists}
                                        onRenderCell={this._onRenderCell}
                                    />
                                </div>
                            }
                            <TextField
                                label='List title'
                                errorMessage={this.state.errorMsg}
                                value={this.state.listTitle}
                                onChange={(_event:React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue:string) => { this.setState({ listTitle: newValue }); this.state.error ? this.setState({ error: false, errorMsg: "" }) : null; }}
                            />
                            <DialogFooter>
                                <PrimaryButton disabled={this.state.error || this.state.listTitle.length <= 0} onClick={this._saveAction} text='Create List' />
                                <DefaultButton onClick={this._closeDialog} text='Cancel' />
                            </DialogFooter>
                        </div>
                    ) : (
                            <div>
                                < Spinner label='Creating list...' />
                            </div>
                        )}
                </Dialog>
            </div>
        );
    }

    private _showDialog = (): void => {
        this.setState({ hideDialog: false });
    }

    private _closeDialog = (): void => {
        this.setState({ hideDialog: true });
    }

    private _saveAction = (): any => {
        const listTitle: string = escape(this.state.listTitle);
        if(!this._validator.checkName(listTitle,ValidationType.ListName)){
            this.setState({ errorMsg: "Invalid list name", error: true });
        }else{
            this.setState({ loading: true });
            return this.props.saveAction(listTitle).then((result: IList) => {
                this.state.createdLists.push(result);
                this.setState({ loading: false, createdLists: this.state.createdLists, listTitle: '' });
            }).catch((error) => {
                console.log(error);
                this.setState({ loading: false, errorMsg: error.message, error: true });
            });
        }
    }
    private _resolveList = (list: List): void => {
        this._list = list;
    }

    private _onRenderCell(item: IList, index: number): JSX.Element {
        return (
            <div className={styles["ms-ListScrollingExample-itemCell"]} data-is-focusable={true}>
                <div
                    className={css(
                        styles["ms-ListScrollingExample-itemContent"],
                        (index % 2 === 0) && styles['ms-ListScrollingExample-itemContent-even'],
                        (index % 2 === 1) && styles['ms-ListScrollingExample-itemContent-odd']
                    )}
                >
                    <Link target="_blank" href={item.NavUrl}> &nbsp; {item.Title}</Link>
                </div>
            </div>
        );
    }
}

