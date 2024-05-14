import * as React from 'react';
import type { ICompleteProjectProps } from './ICompleteProjectProps';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import "@pnp/sp/site-users/web";
import { IBusinesses } from '../models/listInterfaces';
import { useEffect, useState } from 'react';
import { DatePicker, DefaultButton, DirectionalHint, Dropdown, FocusTrapZone, IDropdownOption, IPersonaProps, Icon, Label, Layer, MessageBar, MessageBarType, Popup, PrimaryButton, Stack, StackItem, TextField, TooltipHost, defaultDatePickerStrings } from '@fluentui/react';
import { getSP } from '../../../pnpjsConfig';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react';
import styles from './CompleteProject.module.scss';
import { buttonStackTokens, confirmTitleStyle, dropdown, filePickerButtonUploadStyle, inputStyle, inputStyleLarge, popupStyles, stackTokens, verticalGapStackTokens } from '../constants/styleConstants';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { addItemsToList } from '../services/addItemsToList';
import { countryOptions } from '../constants/optionsConstants';

const CompleteProject = (props: ICompleteProjectProps): JSX.Element => {

  const sp: SPFI = getSP();
  const [businessListItems, setBusinessListItems] = useState<IBusinesses[]>([]);
  const [business, setBusiness] = useState<string | null>('');
  const [selectedBusiness, setSelectedBusiness] = useState<string | null>('');
  const [country, setCountry] = useState<string | null>('');
  const [selectedCountry, setSelectedCountry] = useState<string | null>('');
  const [riskTitle, setRiskTitle] = useState<string>('');
  const [riskDate, setRiskDate] = useState<Date | undefined>();
  const [assignee, setAssignee] = useState<string[]>([]);
  const [peoplePickerSelectedUsers, setPeoplePickerSelectedUsers] = useState<any[]>([]);
  const [notes, setNotes] = useState<string | undefined>('');
  const [riskReport, setRiskReport] = useState<IFilePickerResult[]>();
  const [filledObligatoryComponents, setFilledObligatoryComponents] = useState<String[]>([]);
  const [disableSendButton, setDisableSendButton] = useState<boolean>(true);
  const [isSendButtonClicked, setIsSendButtonClicked] = useState<boolean>(false);
  const [isSaveButtonClicked, setIsSaveButtonClicked] = useState<boolean>(false);
  const [disableSaveButton, setDisableSaveButton] = useState<boolean>(false);
  const [isDeletePopupVisible, setIsDeletePopupVisible] = useState<boolean>(false);
  const [isSavePopupVisible, setIsSavePopupVisible] = useState<boolean>(false);
  const [isSendPopupVisible, setIsSendPopupVisible] = useState<boolean>(false);
  const [currentUserEmail, setCurrentUserEmail] = useState<string>('');
  const [currentUserFullName, setCurrentUserFullName] = useState<string>('');
  const [isSuccess, setIsSuccess] = useState<boolean>(false);


  let _topElement: HTMLElement; //Created this variable to go to the top of webpage

  const getListItems = async (): Promise<IBusinesses[]> => {
    const items = sp.web.lists.getById(props.listGuid).items.orderBy('Title', true)();
    return (await items).map((item) => ({
        id: item.Id,
        title: item.Title,
        country: item.Country,
        client: item.Client
    }));
  }

  useEffect(() => {
    if(props.listGuid && props.listGuid !== '') {
      getListItems().then((items) => {
        setBusinessListItems(items);
      }).catch(error => {
        console.error('Error getting list items:', error);
      });
    }
  }, [props]);

  const getCurrentUser = async () => {
    try {
      const user = await sp.web.currentUser();
      setCurrentUserEmail(user.Email);
      setCurrentUserFullName(user.Title);
    } catch (error) {
      console.error('Error getting current user:', error);
    }
  };
  
  useEffect(() => {
    getCurrentUser();
  }, []);

  const onDropdownChange = (_event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption, name?: string): void => {
    if (item && name === "business") {
      setBusiness(item.text as string);
      setSelectedBusiness(item.key as string);
    }
    else if (item && name === "country") {
      setCountry(item.text as string);
      setSelectedCountry(item.key as string);
    }
    
    if (name && !filledObligatoryComponents.includes(name)) {
      setFilledObligatoryComponents(prevState => [...prevState, name]);
    }

    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsSuccess(false);
  }

  const onTitleChange = (_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value: string | undefined, name: string) => {
    setRiskTitle(value || '');

    if (!filledObligatoryComponents.includes(name)) {
      setFilledObligatoryComponents(prevState => [...prevState, name]);
    }
    else if(filledObligatoryComponents.includes(name) && !value) {
      setFilledObligatoryComponents(prevState => prevState.filter(item => item !== name));
    }

    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsSuccess(false);
  }
  
  const onFormatDate = (date?: Date): string => {
    return (
      !date ? '' 
      : date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear());
  }

  const onDateChange = (value: Date | null | undefined, name: string) => { 
    if(value) {
      setRiskDate(value);
    }

    if (!filledObligatoryComponents.includes(name)) {
      setFilledObligatoryComponents(prevState => [...prevState, name]);
    }

    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsSuccess(false);
  }

  const onPeoplePickerChange = (people: any[], name: string) => {
    const fullNames = people.map(person => person.text);
    setAssignee(fullNames);
    setPeoplePickerSelectedUsers(people);
    
    if (!filledObligatoryComponents.includes(name)) {
      setFilledObligatoryComponents(prevState => [...prevState, name]);
    }
    else if(filledObligatoryComponents.includes(name) && people.length === 0) {
      setFilledObligatoryComponents(prevState => prevState.filter(item => item !== name));
    }

    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsSuccess(false);
  }

  const onUpload = async (files: IFilePickerResult[], name: string) => {
    setRiskReport(files);
    setFilledObligatoryComponents(prevState => [...prevState, name]);
    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsSuccess(false);
  }

  const onDeleteConfirmation = (name: string) => {
    setRiskReport(undefined);
    setFilledObligatoryComponents(prevState => prevState.filter(item => item !== name));
    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsDeletePopupVisible(false);
  }

  const onSubmit = async (state: string) => {

    const assigneeString = assignee.join('\n');

    let states = {
      riskTitle: riskTitle,
      createdBy: currentUserFullName,
      creationDate: new Date(),
      business: business,
      country: country,
      riskDate: riskDate,
      assignee: assigneeString,
      riskReport: riskReport,
      containsDocuments: filledObligatoryComponents.includes("riskReport") ? "Yes, check RiskEventDocumentsList" : "No",
      notes: notes,
      state: state
    };

    await addItemsToList(props.context, states);
  }

  const onSaveConfirmation = () => {
    onSubmit("Modifying");

    setIsSavePopupVisible(false);
    setIsSendButtonClicked(false);
    setDisableSaveButton(false);
    setFilledObligatoryComponents([]);

    // Resetting all variables
    setBusiness(null);
    setSelectedBusiness(null);
    setCountry(null);
    setSelectedCountry(null);
    setRiskTitle('');
    setRiskDate(undefined);
    setAssignee([]);
    setPeoplePickerSelectedUsers([]);
    setNotes('');
    setRiskReport(undefined);
    _topElement.scrollIntoView(); //Using this command to go to the top of webpage
  }

  const onFormValidation = () => {
    setIsSendButtonClicked(true);

    if (filledObligatoryComponents.length === 6) {
      setIsSendPopupVisible(true);
    }
    else{
      _topElement.scrollIntoView(); //Using this command to go to the top of webpage
    }
  }

  const onSendConfirmation = () => {
    onSubmit("Sent");

    setIsSendPopupVisible(false);
    setDisableSendButton(true);
    setIsSuccess(true);
    setIsSendButtonClicked(false);
    setFilledObligatoryComponents([]);

    // Resetting all variables
    setBusiness(null);
    setSelectedBusiness(null);
    setCountry(null);
    setSelectedCountry(null);
    setRiskTitle('');
    setRiskDate(undefined);
    setAssignee([]);
    setPeoplePickerSelectedUsers([]);
    setNotes('');
    setRiskReport(undefined);
    _topElement.scrollIntoView(); //Using this command to go to the top of webpage
  }

  return (
    <div className={styles.completeProject} >

      {/* Created this element to go to the top of webpage */}
      <div ref={(topElement) => _topElement = topElement!} /> 

      {/* SUCCESS SAVE MESSAGE BAR */}
      {
        isSaveButtonClicked === true &&
        <Stack>
          <br /><br />
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            The request has been saved successfully and the state of the request is "Modifying".
          </MessageBar>
        </Stack>
      }

      {/* SUCCESS SEND MESSAGE BAR */}
      {
        isSuccess === true &&
        <Stack>
          <br /><br />
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            The request has been sent successfully and the state of the request is "Sent".
          </MessageBar>
        </Stack>
      }

      {/* ERROR MESSAGE BAR */}
      {
        isSendButtonClicked === true && 
        isSuccess === false && 
        <Stack>
          <br /><br />
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            The program has encountered a problem. You have not filled all of the obligatory fields. 
          </MessageBar>
        </Stack>
      }

      <Stack tokens={verticalGapStackTokens}>

        {/* TITLE DESCRIPTION */}
        <p className={styles.title_description}>This is a form that allows you to submit a new risk event in a specific time period.</p>
        
        <Stack horizontal wrap tokens={stackTokens}>

          {/* BUSINESS */}
          <StackItem>
            <Label 
              className={
                isSendButtonClicked  
                && !filledObligatoryComponents.includes("business")  
                ? styles.errorStyle  
                : undefined
              }
            >
              Select your business *
            </Label>
            <Dropdown
              placeholder="Select an option"
              options={businessListItems.map((item: IBusinesses) => ({
                key: item.id.toString(),
                text: item.title
              }))}
              styles={dropdown}
              onChange={(ev, item) => onDropdownChange(ev, item, "business")}
              selectedKey={selectedBusiness}
            />
          </StackItem>

          {/* COUNTRY */}
          <StackItem>
            <Label 
              className={
                isSendButtonClicked  
                && !filledObligatoryComponents.includes("country")  
                ? styles.errorStyle  
                : undefined
              }
            >
              Select your country *
            </Label>
            <Dropdown
                placeholder="Select an option"
                options={countryOptions}
                styles={dropdown}
                onChange={(ev, item) => onDropdownChange(ev, item, "country")}
                selectedKey={selectedCountry}
              />
          </StackItem>
        </Stack>

        <Stack horizontal wrap tokens={stackTokens}>

          {/* RISK TITLE */}
          <StackItem>
            <Label 
              className={
                isSendButtonClicked  
                && !filledObligatoryComponents.includes("riskTitle")  
                ? styles.errorStyle  
                : undefined
              }
            >
              Risk event title *
            </Label>
            <TextField 
              className={styles.componentStyle} 
              placeholder="Please write your text here"
              value={riskTitle}
              onChange={(ev, newValue) => onTitleChange(ev, newValue, "riskTitle")}
            />
          </StackItem>

          {/* RISK DATE */}
          <StackItem>
            <Label 
              className={
                isSendButtonClicked  
                && !filledObligatoryComponents.includes("riskDate")  
                ? styles.errorStyle  
                : undefined
              }
            >
              Risk event date *
            </Label>
            <DatePicker
              placeholder="Select a date"
              ariaLabel="Select a date"
              strings={defaultDatePickerStrings}
              formatDate={onFormatDate}
              styles={inputStyle}
              value={riskDate}
              onSelectDate={(newValue) => onDateChange(newValue, "riskDate")}
            />
          </StackItem>
        </Stack>

        <Stack tokens={stackTokens}>

          {/* RISK ASSIGNMENT */}
          <Stack horizontal>
            <Label 
              className={
                isSendButtonClicked  
                && !filledObligatoryComponents.includes("riskAssignment")  
                ? styles.errorStyle  
                : undefined
              }
            >
              Assign to *
            </Label>
            <TooltipHost 
              content="Write at least 3 letters for the name to appear and you can assign upto 3 people."
              directionalHint={DirectionalHint.topLeftEdge}
            >
              <Icon iconName="Info" className={styles.peoplePickerIcon} />
            </TooltipHost>
          </Stack>
          <Stack style={{ marginTop: 0 }}>
            <PeoplePicker
              context={props.context}
              placeholder="Please write the name here"
              personSelectionLimit={3}
              groupName={""} 
              required={false}
              disabled={false}
              searchTextLimit={2}
              onChange={(newPeople) => onPeoplePickerChange(newPeople, "riskAssignment")}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000} 
              styles={inputStyleLarge}
              defaultSelectedUsers={peoplePickerSelectedUsers}
              resultFilter={(result: IPersonaProps[]) => {
                return result.filter(person => person.secondaryText !== currentUserEmail);
              }}
            />
          </Stack>
        </Stack>

        <Stack horizontal wrap tokens={stackTokens}>

          {/* NOTES */}
          <StackItem>
            <TextField 
              className={styles.componentStyle} 
              label="Notes" 
              placeholder="Please write your text here"
              multiline 
              rows={5} 
              value={notes}
              onChange={(_ev, newValue) => {
                setNotes(newValue || '');
                setIsSaveButtonClicked(false);
                setIsSendButtonClicked(false);
              }}
            />
          </StackItem>

          {/* RISK REPORT */}
          <StackItem styles={inputStyle}>
            <Label 
              className={
                isSendButtonClicked  
                && !filledObligatoryComponents.includes("riskReport")  
                ? styles.errorStyle  
                : undefined
              }
            >
              Risk event report *
            </Label>
            <Stack className={styles.reportContainerZone}>
              <Stack className={styles.reportContentZone}>
                {riskReport === null || riskReport === undefined ?
                  <p className={styles.filePickerDescription}>Upload files from your local device using the button below</p>
                  :
                  <p className={styles.filePickerDescription}>If you want you can delete your files using the button below</p>
                }
                <Stack horizontal verticalAlign="center">
                    {riskReport === null || riskReport === undefined ?
                      <FilePicker
                        context={props.context as any}
                        accepts={[ ".pdf" ]}
                        hidden={false}
                        hideLocalUploadTab={true}
                        hideLocalMultipleUploadTab={false}
                        hideOneDriveTab={true}
                        hideStockImages={true}
                        hideWebSearchTab={true}
                        hideSiteFilesTab={true}
                        hideLinkUploadTab={true}
                        hideRecentTab={true}
                        onSave={(ev) => onUpload(ev, "riskReport")}
                        buttonIconProps={{ styles: filePickerButtonUploadStyle }}
                        buttonClassName={styles['filePickerButtonUpload']}
                        buttonLabel='Upload'
                      />
                      :
                      <PrimaryButton 
                        text="Delete" 
                        onClick={() => setIsDeletePopupVisible(true)}
                      />
                    }
                    <StackItem>
                      {riskReport === null || riskReport === undefined ?
                        <p className={styles.filePickerFormat}>pdf formats allowed</p>
                        :
                        <></>
                      }
                    </StackItem>
                </Stack>
              </Stack>
            </Stack>
          </StackItem>
        </Stack>

        {/* BOTTOM SECTION */}
        <Stack horizontal horizontalAlign="start" className={styles.bottomSection}>
          <StackItem>
            <Stack horizontal tokens={buttonStackTokens}>
              <DefaultButton 
                text="Save to draft" 
                disabled={disableSaveButton} 
                onClick={() => {
                  setIsSavePopupVisible(true);
                  setIsSaveButtonClicked(true);
                }} 
              />
              <PrimaryButton 
                text="Finish and Send" 
                disabled={filledObligatoryComponents.length === 0 && disableSendButton} 
                onClick={onFormValidation} 
              />
              <p className={styles.obligatoryField}>* Obligatory field</p>
            </Stack>
          </StackItem>
        </Stack>
      
      </Stack>

      {/* POPUP SECTION */}
      {/* DELETE BUTTON POPUP */}
      {
        isDeletePopupVisible === true && (
          <Layer>
            <Popup
              className={popupStyles.root}
              role="dialog"
              onDismiss={() => setIsDeletePopupVisible(false)}
            >
              <FocusTrapZone>
                <div className={popupStyles.content}>
                  <p style={confirmTitleStyle}>Are you sure you want to delete your uploaded files?</p>
                  <p>By clicking "Yes" your uploaded files will be deleted and you can upload new files.</p>
                  <Stack horizontal horizontalAlign="center">
                    <DefaultButton 
                      className={styles.popupDefaultButton} 
                      onClick={() => setIsDeletePopupVisible(false)}
                    >
                      No
                    </DefaultButton>
                    <PrimaryButton 
                      className={styles.popupPrimaryButton} 
                      onClick={() => onDeleteConfirmation("riskReport")} 
                    >
                      Yes
                    </PrimaryButton>
                  </Stack>
                </div>
              </FocusTrapZone>
            </Popup>
          </Layer>
        )
      }

      {/* SAVE BUTTON POPUP */}
      {
        isSavePopupVisible === true && (
          <Layer>
            <Popup
              className={popupStyles.root}
              role="dialog"
              onDismiss={() => {
                setIsSavePopupVisible(false); 
                setIsSaveButtonClicked(false);
                setIsSendButtonClicked(false);
              }}
            >
              <FocusTrapZone>
                <div className={popupStyles.content}>
                  <p style={confirmTitleStyle}>Are you sure you want to save your session to draft?</p>
                  <p>By clicking "Yes" your session will be saved as draft and in modifying state.</p>
                  <Stack horizontal horizontalAlign="center">
                    <DefaultButton 
                      className={styles.popupDefaultButton} 
                      onClick={() => {
                        setIsSavePopupVisible(false); 
                        setIsSaveButtonClicked(false);
                        setIsSendButtonClicked(false);
                      }}
                    >
                      No
                    </DefaultButton>
                    <PrimaryButton 
                      className={styles.popupPrimaryButton} 
                      onClick={() => onSaveConfirmation()} 
                    >
                      Yes
                    </PrimaryButton>
                  </Stack>
                </div>
              </FocusTrapZone>
            </Popup>
          </Layer>
        )
      }

      {/* SEND BUTTON POPUP */}
      {
        isSendPopupVisible === true && (
          <Layer>
            <Popup
              className={popupStyles.root}
              role="dialog"
              onDismiss={() => {
                setIsSendPopupVisible(false);
                setIsSendButtonClicked(false);
              }}
            >
              <FocusTrapZone>
                <div className={popupStyles.content}>
                  <p style={confirmTitleStyle}>Are you sure you want to finish your session and send it?</p>
                  <p>By clicking "Yes" your session will be completed and in sent state.</p>
                  <Stack horizontal horizontalAlign="center">
                    <DefaultButton 
                      className={styles.popupDefaultButton} 
                      onClick={() => {
                        setIsSendPopupVisible(false);
                        setIsSendButtonClicked(false);
                      }}
                    >
                      No
                    </DefaultButton>
                    <PrimaryButton 
                      className={styles.popupPrimaryButton} 
                      onClick={() => onSendConfirmation()} 
                    >
                      Yes
                    </PrimaryButton>
                  </Stack>
                </div>
              </FocusTrapZone>
            </Popup>
          </Layer>
        )
      }
    </div>
  )
}

export default CompleteProject;