import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import WebViewer, { UI, WebViewerInstance } from '@pdftron/webviewer';
import type { Core } from '@pdftron/webviewer';

import * as strings from 'WebviewerWebPartStrings';

export interface IWebviewerWebPartProps {
  description: string;
}

type AccessMode = 'read' | 'full';
type DemoUserRole = 'admin' | 'reviewerA' | 'reviewerB' | 'default';

interface ISharePointBasePermissions {
  High?: string | number;
  Low?: string | number;
}

export default class WebviewerWebPart extends BaseClientSideWebPart<IWebviewerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _mode: string;
  private _accessMode: AccessMode = 'full';
  private _modalMessageElement: HTMLElement;
  private _webViewerInstance: WebViewerInstance | undefined;
  private _viewerContainer: HTMLElement | undefined;
  private _viewerInitKey: string = '';
  private _renderGeneration: number = 0;
  private _demoUserRole: DemoUserRole = 'default';
  private _canPersistChanges: boolean = true;

  public validateQueryParam(urlParams: URLSearchParams): boolean {
    return !!urlParams.get('fileUrl') || (!!urlParams.get('filename') && !!urlParams.get('foldername'));
  }

  public render(): void {
    this.domElement.style.height = '1000px';
    const siteRelativeUrl: string = this._siteServerRelativeUrl();
    const sampleFileServerRelativeUrl: string = `${siteRelativeUrl}/${process.env.FOLDER_URL}/webviewer-sharepoint-sample.pdf`;
    const urlParams: URLSearchParams = new URLSearchParams(window.location.search);
    const requestedFileServerRelativeUrl: string = this._getRequestedFileServerRelativeUrl(urlParams, siteRelativeUrl);
    const fileServerRelativeUrl: string = requestedFileServerRelativeUrl || sampleFileServerRelativeUrl;
    const initialFileName: string = urlParams.get('filename') || this._getFileNameFromServerRelativeUrl(fileServerRelativeUrl);
    const initialDocUrl: string = `${window.location.origin}${siteRelativeUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${this._escapeODataString(fileServerRelativeUrl)}')/$value`;
    const viewerInitKey: string = `${fileServerRelativeUrl}|${initialFileName}`;

    if (this._viewerInitKey === viewerInitKey && this._viewerContainer && this.domElement.contains(this._viewerContainer)) {
      return;
    }

    this._renderGeneration++;
    const renderGeneration: number = this._renderGeneration;
    this._viewerInitKey = viewerInitKey;
    this._disposeWebViewerInstance();
    this.domElement.replaceChildren();

    this._viewerContainer = document.createElement('div');
    this._viewerContainer.style.height = '100%';
    this._viewerContainer.style.width = '100%';
    this.domElement.appendChild(this._viewerContainer);

    WebViewer({
      // We suggest to use the method of uploading static files to the Documents folder in your sharepoint site
      // The provided path below is a template, it may varies in your site
      path: `https://${process.env.TENANT_ID}.sharepoint.com/sites/${process.env.SITE_NAME}/Shared%20Documents/${process.env.WEBVIEWER_LIB_FOLDER_PATH}`,
      // SharePoint Online's CSP does not allow script-src blob:, so force WebViewer's PDF worker
      // to load its worker JavaScript files directly instead of wrapping them in object URL blobs.
      disableObjectURLBlobs: true,
    }, this._viewerContainer)
    .then(async instance => {
      if (renderGeneration !== this._renderGeneration) {
        this._disposeWebViewerInstance(instance);
        return;
      }

      this._webViewerInstance = instance;
      // SharePoint Online's CSP blocks inline scripts. WebViewer's embedded PDF JavaScript
      // support uses an iframe with inline scripts for AcroForm actions, so disable it before
      // loading documents in this SharePoint-hosted sample.
      instance.Core.disableEmbeddedJavaScript();

      this._accessMode = await this._resolveAccessMode(fileServerRelativeUrl);
      const currentUserName: string = this._getCurrentUserEmail() || this.context.pageContext.user.displayName || this.context.pageContext.user.loginName;
      const userData: UI.MentionsManager.UserData[] = [{
        value: currentUserName,
        email: this.context.pageContext.user.email
      }];
      instance.UI.mentions.setUserData(userData);
      instance.Core.annotationManager.setCurrentUser(currentUserName);

      const { Feature } = instance.UI;
      instance.UI.enableFeatures([Feature.FilePicker]);
      const validateQueryParamResult: boolean = this.validateQueryParam(urlParams);
      if (validateQueryParamResult) {
        this._mode = "sharepoint-file";
      } else {
        this._mode = "local-file";
      }

      this._createSavedModal(instance);
      this._createMessageModal(instance);
      this._installAnnotationPermissionPolicy(instance, currentUserName);
      this._applyAccessMode(instance, this._accessMode);
      this._showWelcomeMessageAfterDocumentLoad(instance, renderGeneration);
      instance.UI.loadDocument(initialDocUrl, { filename: initialFileName });
    })
    .catch(err => console.error(err));
  }

  private _disposeWebViewerInstance(instance: WebViewerInstance = this._webViewerInstance): void {
    if (!instance) {
      return;
    }

    interface IDisposableUI {
      dispose?: () => Promise<void>;
    }

    try {
      const disposePromise: Promise<void> | undefined = (instance.UI as unknown as IDisposableUI).dispose?.();
      disposePromise?.catch(error => console.warn('Unable to dispose existing WebViewer instance.', error));
    } catch (error) {
      console.warn('Unable to dispose existing WebViewer instance.', error);
    }

    if (instance === this._webViewerInstance) {
      this._webViewerInstance = undefined;
    }
  }

  private _siteServerRelativeUrl(): string {
    return `/sites/${process.env.SITE_NAME}`;
  }

  private _escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }

  private _getRequestedFileServerRelativeUrl(urlParams: URLSearchParams, siteRelativeUrl: string): string | undefined {
    const fileUrl: string = urlParams.get('fileUrl');
    if (fileUrl) {
      return fileUrl;
    }

    const filename: string = urlParams.get('filename');
    const folderName: string = urlParams.get('foldername');
    if (filename && folderName) {
      return `${siteRelativeUrl}/${folderName}/${filename}`;
    }

    return undefined;
  }

  private _getFileNameFromServerRelativeUrl(fileServerRelativeUrl: string): string {
    const pathParts: string[] = fileServerRelativeUrl.split('/');
    return pathParts[pathParts.length - 1];
  }

  private _getFolderUrlFromServerRelativeUrl(fileServerRelativeUrl: string): string {
    return fileServerRelativeUrl.substring(0, fileServerRelativeUrl.lastIndexOf('/'));
  }

  private _normalizeIdentity(value: string): string {
    return (value || '').toLowerCase().replace(/^.*\|/, '').trim();
  }

  private _getCurrentUserEmail(): string {
    const email: string = this.context.pageContext.user.email || '';
    if (email) {
      return this._normalizeIdentity(email);
    }

    return this._normalizeIdentity(this.context.pageContext.user.loginName || '');
  }

  private _resolveDemoUserRole(): DemoUserRole {
    const userIdentity: string = `${this.context.pageContext.user.email || ''} ${this.context.pageContext.user.loginName || ''}`.toLowerCase();
    if (userIdentity.indexOf('yixiaochen@yctestio.onmicrosoft.com') >= 0) {
      return 'admin';
    }
    if (userIdentity.indexOf('gao@yctestio.onmicrosoft.com') >= 0) {
      return 'reviewerA';
    }
    if (userIdentity.indexOf('unauthorized@yctestio.onmicrosoft.com') >= 0) {
      return 'reviewerB';
    }

    return 'default';
  }

  private async _resolveAccessMode(fileServerRelativeUrl: string): Promise<AccessMode> {
    this._demoUserRole = this._resolveDemoUserRole();

    if (this._demoUserRole === 'admin' || this._demoUserRole === 'reviewerA' || this._demoUserRole === 'reviewerB') {
      this._canPersistChanges = true;
      return 'full';
    }

    try {
      const permissions: ISharePointBasePermissions = await this._getFileEffectiveBasePermissions(fileServerRelativeUrl);
      const canEdit: boolean = this._canEditListItems(permissions);
      this._canPersistChanges = canEdit;
      return canEdit ? 'full' : 'read';
    } catch (error) {
      console.warn('Unable to resolve SharePoint permissions. Falling back to read-only mode.', error);
      return 'read';
    }
  }

  private async _getFileEffectiveBasePermissions(fileServerRelativeUrl: string): Promise<ISharePointBasePermissions> {
    const siteRelativeUrl: string = this._siteServerRelativeUrl();
    const resp: Response = await fetch(`${window.location.origin}${siteRelativeUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${this._escapeODataString(fileServerRelativeUrl)}')/ListItemAllFields/effectiveBasePermissions`, {
      method: 'GET',
      credentials: 'same-origin',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    });

    if (!resp.ok) {
      throw new Error(`SharePoint permissions request failed: ${resp.status} ${resp.statusText}`);
    }

    const responseJson: unknown = await resp.json();
    const json: { d?: unknown; EffectiveBasePermissions?: unknown; High?: string | number; Low?: string | number } = responseJson as { d?: unknown; EffectiveBasePermissions?: unknown; High?: string | number; Low?: string | number };
    const d: { EffectiveBasePermissions?: unknown; High?: string | number; Low?: string | number } = json.d as { EffectiveBasePermissions?: unknown; High?: string | number; Low?: string | number };
    return (json.EffectiveBasePermissions || d?.EffectiveBasePermissions || d || json) as ISharePointBasePermissions;
  }

  private _canEditListItems(permissions: ISharePointBasePermissions): boolean {
    const lowPermissions: number = typeof permissions.Low === 'string' ? parseInt(permissions.Low, 10) : permissions.Low || 0;
    const editListItemsPermission: number = 4;
    return (lowPermissions & editListItemsPermission) === editListItemsPermission;
  }

  private _applyAccessMode(instance: WebViewerInstance, accessMode: AccessMode): void {
    if (accessMode === 'read') {
      instance.UI.enableViewOnlyMode();
      instance.UI.disableElements(['saveFileButton']);
      this._installReadOnlyAnnotationGuard(instance);
      return;
    }

    this._createSaveFileButton(instance);
  }

  private _showWelcomeMessage(instance: WebViewerInstance): void {
    let welcomeMessage: { title: string; message: string };
    if (this._demoUserRole === 'admin') {
      welcomeMessage = {
        title: 'Welcome Admin',
        message: 'You can review, annotate, and manage all reviewers\' annotations.'
      };
    } else if (this._demoUserRole === 'reviewerA') {
      welcomeMessage = {
        title: 'Welcome Reviewer A',
        message: 'You can add and manage your own annotations.'
      };
    } else if (this._demoUserRole === 'reviewerB') {
      welcomeMessage = {
        title: 'Welcome Reviewer B',
        message: 'You can add annotations, but cannot modify annotations created by other reviewers.'
      };
    } else {
      welcomeMessage = {
        title: 'Welcome',
        message: this._accessMode === 'read' ? 'You are in read only mode' : 'You can review and annotate the document.'
      };
    }

    instance.UI.showWarningMessage({
      title: welcomeMessage.title,
      message: welcomeMessage.message,
      confirmBtnText: 'OK',
      onConfirm: () => Promise.resolve(),
      onCancel: () => Promise.resolve()
    });
  }

  private _showWelcomeMessageAfterDocumentLoad(instance: WebViewerInstance, renderGeneration: number): void {
    let welcomeMessageShown: boolean = false;
    instance.Core.documentViewer.addEventListener('documentLoaded', () => {
      if (welcomeMessageShown || renderGeneration !== this._renderGeneration) {
        return;
      }

      welcomeMessageShown = true;
      window.setTimeout(() => {
        if (renderGeneration === this._renderGeneration) {
          this._showWelcomeMessage(instance);
        }
      }, 250);
    });
  }

  private _installAnnotationPermissionPolicy(instance: WebViewerInstance, currentUserName: string): void {
    const annotationManager: Core.AnnotationManager = instance.Core.annotationManager;
    const documentViewer: Core.DocumentViewer = instance.Core.documentViewer;
    const annotationHistoryManager: Core.AnnotationHistoryManager = documentViewer.getAnnotationHistoryManager();
    let baselineXfdf: string = '';
    let reverting: boolean = false;

    annotationManager.setPermissionCheckCallback((author: string, annotation: Core.Annotations.Annotation) => {
      if (this._demoUserRole === 'admin') {
        return true;
      }
      if (this._accessMode === 'read') {
        return false;
      }

      const annotationAuthor: string = this._normalizeIdentity(annotation?.Author || author || '');
      return !annotationAuthor || annotationAuthor === currentUserName;
    });

    documentViewer.addEventListener('documentLoaded', async () => {
      baselineXfdf = await annotationManager.exportAnnotations();
    });

    annotationManager.addEventListener('annotationChanged', (annotations: Core.Annotations.Annotation[], action: string, info: { imported?: boolean; isUndoRedo?: boolean }) => {
      if (this._accessMode === 'read' || reverting || info?.imported || info?.isUndoRedo) {
        return;
      }

      if (this._demoUserRole === 'admin') {
        window.setTimeout(async () => {
          baselineXfdf = await annotationManager.exportAnnotations();
        }, 0);
        return;
      }

      const hasOwnershipViolation: boolean = annotations.some((annotation: Core.Annotations.Annotation) => {
        const annotationAuthor: string = this._normalizeIdentity(annotation.Author || '');
        return !!annotationAuthor && annotationAuthor !== currentUserName;
      });

      if (!hasOwnershipViolation || (action !== 'modify' && action !== 'delete')) {
        window.setTimeout(async () => {
          baselineXfdf = await annotationManager.exportAnnotations();
        }, 0);
        return;
      }

      reverting = true;
      window.setTimeout(async () => {
        try {
          if (annotationHistoryManager.canUndo()) {
            await annotationHistoryManager.undo();
          } else if (baselineXfdf) {
            await annotationManager.importAnnotations(baselineXfdf);
          }

          instance.UI.displayErrorMessage('Security Restriction: You are not authorized to modify or delete annotations created by another reviewer.');
        } finally {
          reverting = false;
        }
      }, 1000);
    });
  }

  private _installReadOnlyAnnotationGuard(instance: WebViewerInstance): void {
    const annotationManager: Core.AnnotationManager = instance.Core.annotationManager;
    const documentViewer: Core.DocumentViewer = instance.Core.documentViewer;
    let baselineXfdf: string = '';
    let restoring: boolean = false;

    const markAnnotationsReadOnly = (): void => {
      annotationManager.getAnnotationsList().forEach((annotation: Core.Annotations.Annotation) => {
        annotation.ReadOnly = true;
      });
      annotationManager.drawAnnotationsFromList(annotationManager.getAnnotationsList());
    };

    documentViewer.addEventListener('documentLoaded', async () => {
      markAnnotationsReadOnly();
      baselineXfdf = await annotationManager.exportAnnotations();
    });

    annotationManager.addEventListener('updateAnnotationPermission', (annotation?: Core.Annotations.Annotation) => {
      if (annotation) {
        annotation.ReadOnly = true;
        return;
      }

      markAnnotationsReadOnly();
    });

    annotationManager.addEventListener('annotationChanged', async (annotations: Core.Annotations.Annotation[], action: string, info: { imported?: boolean; isUndoRedo?: boolean }) => {
      if (restoring || info?.imported || info?.isUndoRedo) {
        return;
      }

      restoring = true;
      try {
        if (action === 'add') {
          annotationManager.deleteAnnotations(annotations, { imported: true, force: true });
        } else if (baselineXfdf) {
          await annotationManager.importAnnotations(baselineXfdf);
          markAnnotationsReadOnly();
        }

        this._showMessage(instance, 'Change blocked', 'Your current role is read-only. The attempted annotation change was reverted.');
      } finally {
        restoring = false;
      }
    });
  }

  private _createSaveFileButton(instance: WebViewerInstance): void {
    const saveFile = async (): Promise<void> => {
      if (!this._canPersistChanges) {
        instance.UI.displayErrorMessage('Security Violation: Unauthorized document save detected. The document was not saved.');
        return;
      }

      instance.UI.openElements(['loadingModal']);
      try {
        if (this._mode === 'sharepoint-file') {
          const searchparams: URLSearchParams = new URLSearchParams(window.location.search);
          const fileUrl: string = searchparams.get('fileUrl');
          const folderName: string = fileUrl ? this._getFolderUrlFromServerRelativeUrl(fileUrl) : searchparams.get('foldername');
          const fileName: string = fileUrl ? this._getFileNameFromServerRelativeUrl(fileUrl) : searchparams.get('filename');
          await this.saveFile(instance, folderName, fileName);
        } else if (this._mode === 'local-file') {
          const fileName: string = await instance.Core.documentViewer.getDocument().getFilename();
          const folderName: string = encodeURIComponent(process.env.FOLDER_URL);
          await this.saveFile(instance, folderName, fileName);
        }
        instance.UI.openElements(['savedModal']);
      } catch (error) {
        console.error(error);
        this._showMessage(instance, 'Save failed', 'SharePoint rejected the save request. Check your file permissions and try again.');
      } finally {
        instance.UI.closeElements(['loadingModal']);
      }
    };

    interface IModularHeader {
      items: unknown[];
      setItems: (items: unknown[]) => void;
    }

    interface IModularUI {
      Components?: {
        CustomButton?: new (options: unknown) => unknown;
      };
      getModularHeader?: (dataElement: string) => IModularHeader;
    }

    const modularUI: IModularUI = instance.UI as unknown as IModularUI;
    const defaultHeader: IModularHeader = modularUI.getModularHeader?.('default-top-header');
    if (modularUI.Components?.CustomButton && defaultHeader) {
      const saveFileButton: unknown = new modularUI.Components.CustomButton({
        dataElement: 'saveFileButton',
        className: 'save-file-button',
        label: 'Save',
        title: 'Save file to SharePoint',
        onClick: saveFile,
        img: 'icon-save',
        style: {
          backgroundColor: '#F1F3F5'
        }
      });
      const existingItems: unknown[] = defaultHeader.items || [];
      const saveButtonExists: boolean = existingItems.some((item: { dataElement?: string } | string) => item === 'saveFileButton' || (typeof item !== 'string' && item?.dataElement === 'saveFileButton'));
      if (!saveButtonExists) {
        defaultHeader.setItems([...existingItems, saveFileButton]);
      }
      return;
    }

    instance.UI.setHeaderItems((header: UI.Header) => {
      const saveFileButton: unknown = {
        type: 'actionButton',
        dataElement: 'saveFileButton',
        title: 'Save file to SharePoint',
        img: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M0 0h24v24H0z" fill="none"/><path d="M17 3H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V7l-4-4zm-5 16c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H5V5h10v4z"/></svg>',
        onClick: saveFile
      };
      header.get('view-controls').insertBefore(saveFileButton);
    });
  }

  /* 
    The purpose of this function is to get the request digest (client-side token) for us to go through the authorization
    when uploading the file.
  */
  private async _getFormDigestValue(): Promise<string> {
    try {
      const resp: Response = await fetch(`${window.location.origin}/sites/${process.env.SITE_NAME}/_api/contextinfo`, {
        method: 'POST',
        headers: {
          'Accept': 'application/json; odata=verbose'
        },
      });
      
      interface IDigestResponseJson {
        d: {
          GetContextWebInformation: {
            FormDigestValue: string
          }
        }
      }
      const respJson: IDigestResponseJson = await resp.json();
      return respJson.d.GetContextWebInformation.FormDigestValue;
    } catch(error) {
      console.error(error);
    }
  }

  public async saveFile(instance: WebViewerInstance, folderUrl: string, fileName: string): Promise<void> {
    if (!this._canPersistChanges) {
      throw new Error('Current user is not authorized to save the document.');
    }

    const annotationManager: Core.AnnotationManager = instance.Core.annotationManager;
    const xfdfString: string = await annotationManager.exportAnnotations();
    const fileData: ArrayBuffer = await instance.Core.documentViewer.getDocument().getFileData({ xfdfString });
    const digest: string = await this._getFormDigestValue();
    const fileBlob: Blob = new Blob([fileData], {
      type: 'application/pdf'
    });
    const file: File = new File([fileBlob], fileName, {
      type: 'application/pdf'
    });
    const siteRelativeUrl: string = this._siteServerRelativeUrl();
    const folderServerRelativeUrl: string = folderUrl.startsWith('/') ? folderUrl : `${siteRelativeUrl}/${folderUrl}`;
    const resp: Response = await fetch(`${window.location.origin}${siteRelativeUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${this._escapeODataString(folderServerRelativeUrl)}')/Files/add(url='${this._escapeODataString(fileName)}', overwrite=true)`, {
      method: 'POST',
      body: file,
      headers: {
        'accept': 'application/json; odata=verbose',
        'X-RequestDigest': digest,
        'Content-Length': fileData.byteLength.toString()
      }
    });

    if (!resp.ok) {
      throw new Error(`SharePoint save failed: ${resp.status} ${resp.statusText}`);
    }
  }

  private _createSavedModal(instance: WebViewerInstance): void {
    const divInput: HTMLElement = document.createElement('div');
    divInput.innerText = 'File saved successfully';

    interface IModal { 
      dataElement: string;
      disableBackdropClick?: boolean; 
      disableEscapeKeyDown?: boolean; 
      render: UI.renderCustomModal; 
      header: unknown; 
      body: unknown; 
      footer: unknown; 
    }

    const modal: IModal = {
      dataElement: 'savedModal',
      body: {
        className: 'myCustomModal-body',
        style: {
          'text-align': 'center'
        },
        children: [divInput]
      },
      header: null,
      footer: null,
      render: null
    }
    instance.UI.addCustomModal(modal);
  }

  private _createMessageModal(instance: WebViewerInstance): void {
    this._modalMessageElement = document.createElement('div');
    this._modalMessageElement.innerText = '';

    interface IModal { 
      dataElement: string;
      disableBackdropClick?: boolean; 
      disableEscapeKeyDown?: boolean; 
      render: UI.renderCustomModal; 
      header: unknown; 
      body: unknown; 
      footer: unknown; 
    }

    const modal: IModal = {
      dataElement: 'accessMessageModal',
      body: {
        className: 'accessMessageModal-body',
        style: {
          'text-align': 'center'
        },
        children: [this._modalMessageElement]
      },
      header: null,
      footer: null,
      render: null
    };
    instance.UI.addCustomModal(modal);
  }

  private _showMessage(instance: WebViewerInstance, title: string, message: string): void {
    if (this._modalMessageElement) {
      this._modalMessageElement.innerText = `${title}\n\n${message}`;
    }

    instance.UI.openElements(['accessMessageModal']);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  protected onDispose(): void {
    this._renderGeneration++;
    this._disposeWebViewerInstance();
    this._viewerContainer = undefined;
    this.domElement.replaceChildren();
  }



  private _getEnvironmentMessage(): string {
    if (this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
