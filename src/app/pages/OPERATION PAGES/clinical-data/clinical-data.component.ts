import {
  Component,
  ElementRef,
  NgModule,
  OnInit,
  ViewChild,
} from '@angular/core';
import notify from 'devextreme/ui/notify';
import {
  DxButtonModule,
  DxDataGridComponent,
  DxDataGridModule,
  DxDateBoxModule,
  DxDropDownBoxModule,
  DxDropDownButtonModule,
  DxLoadPanelModule,
  DxLookupModule,
  DxPopupModule,
  DxProgressBarModule,
  DxSelectBoxModule,
  DxTagBoxModule,
  DxTextAreaModule,
  DxTextBoxModule,
  DxAccordionModule,
} from 'devextreme-angular';
import { ReportService } from 'src/app/services/Report-data.service';
import { CommonModule } from '@angular/common';
import { FormPopupModule } from 'src/app/components';
import DataSource from 'devextreme/data/data_source';
import { Router } from '@angular/router';
import { DataService } from 'src/app/services';
import { OperationReportService } from '../operation-report.service';
import { ClinicalDataImportFormModule } from '../../POP-UP_PAGES/clinical-data-import-form/clinical-data-import-form.component';
import { DatePipe } from '@angular/common';
import { firstValueFrom } from 'rxjs';
import { MasterReportService } from '../../MASTER PAGES/master-report.service';
import { NotificationService } from 'src/app/services/notification.service';

@Component({
  selector: 'app-clinical-data',
  templateUrl: './clinical-data.component.html',
  styleUrls: ['./clinical-data.component.scss'],
  providers: [ReportService, DataService, OperationReportService, DatePipe],
})
export class ClinicalDataComponent implements OnInit {
  @ViewChild(DxDataGridComponent, { static: true })
  dataGrid!: DxDataGridComponent;

  @ViewChild('fileInput')
  fileInputRef!: ElementRef<HTMLInputElement>;

  // Pagination
  readonly allowedPageSizes = [5, 10, 'all'];
  displayMode = 'full';
  showPageSizeSelector = true;
  showInfo = true;
  showNavButtons = true;

  // Grid
  isFilterRowVisible = false;
  focusedRowKey: any = null;

  importButtonOptions = {
    icon: 'import',
    type: 'default',
    stylingMode: 'contained',
    hint: 'Import Xml Files',
    onClick: () => {
      if (!this.selectedFacility) {
        this.notificationService.showNotification(
          'Please select a facility before importing XML files.',
          'error',
        );
        return;
      }

      this.fileInputRef.nativeElement.click();
    },
    elementAttr: { class: 'add-button' },
  };

  processButtonOptions = {
    icon: '',
    text: 'Upload',
    type: 'default',
    stylingMode: 'contained',
    hint: 'Upload Selected Data',
    onClick: () => this.uploadXmlFile(),
    elementAttr: { class: 'add-button' },
  };

  // Facility
  facilityListDataSource: any[] = [];
  selectedFacility: any[] = [];

  // Filters
  fromDate: any = null;
  toDate: any = null;
  today: Date = new Date();

  selectedmonth: any = '';
  selectedYear: any = null;

  minDate: Date;
  maxDate: Date;

  monthDataSource: { name: string; value: any }[];
  years: number[] = [];

  // XML Import
  isExcelLoading = false;
  isResponsePopupOpened = false;
  hasError = false;

  importResults: any[] = [];
  selectedXmlFile: any[] = [];

  failCount = 0;
  alreadyImportedCount = 0;
  successCount = 0;
  uploadedCount = 0;
  totalFiles = 0;

  // Grid Selection
  selectedRowKeys: any[] = [];
  selectedRowsData: any[] = [];

  // Upload
  uploadOptions = [
    { text: 'Test Upload', value: true },
    { text: 'Production Upload', value: false },
  ];

  uploadMode = true;

  isUploading = false;
  currentRequest: any;
  isCancelled = false;

  totalCount = 0;
  completedCount = 0;
  failedCount = 0;

  // XML Data
  xmlGridData: any;
  showXmlPopup = false;

  uploadedFileXMlData: any;

  selectedXmlRow: any;
  xmlData: string = '';
  XmlFileName: string = '';
  clickedrowData: any;

  showXmlDetailsPopup = false;
  loadingVisible = false;

  // Process ID
  processID: any = 0;

  constructor(
    private service: ReportService,
    private router: Router,
    private dataService: DataService,
    private operationService: OperationReportService,
    private datePipe: DatePipe,
    private masterService: MasterReportService,
    private notificationService: NotificationService,
  ) {
    this.minDate = new Date(2023, 0, 1);
    this.maxDate = new Date(); // Set the maximum date
    //============Year field dataSource===============
    const currentYear = new Date().getFullYear();
    for (let year = currentYear; year >= 2023; year--) {
      this.years.push(year);
    }
    //=============month field datasource============
    this.monthDataSource = this.service.getMonths();
  }

  async ngOnInit() {
    try {
      await this.loadFacilityData();
      this.selectedYear = 2026;
      const today = new Date();
      this.toDate = today;
      this.fromDate = new Date(today.getFullYear(), 0, 1);
      this.isFilterRowVisible = false;
      this.onApplyFilter();
    } catch (error) {
      console.error('Initialization error:', error);
    }
  }

  //================ Year value change ===================
  onYearChanged(e: any): void {
    this.selectedYear = e.value;
    this.selectedmonth = '';
    const currentYear = new Date().getFullYear();
    const today = new Date();
    if (this.selectedYear === currentYear) {
      // Set from date to the start of the year and to date to today
      this.fromDate = new Date(this.selectedYear, 0, 1); // January 1 of the current year
      this.toDate = today; // Today's date
    } else {
      this.fromDate = new Date(this.selectedYear, 0, 1); // January 1
      this.toDate = new Date(this.selectedYear, 11, 31); // December 31
    }
  }

  //================Month value change ===================
  onMonthValueChanged(e: any) {
    this.selectedmonth = e.value ?? '';

    const today = new Date();
    const currentYear = today.getFullYear();

    if (this.selectedmonth === '') {
      if (this.selectedYear === currentYear) {
        this.fromDate = new Date(currentYear, 0, 1);
        this.toDate = today;
      } else {
        this.fromDate = new Date(this.selectedYear, 0, 1);
        this.toDate = new Date(this.selectedYear, 11, 31);
      }
    } else {
      this.fromDate = new Date(this.selectedYear, this.selectedmonth, 1);
      this.toDate = new Date(this.selectedYear, this.selectedmonth + 1, 0);
    }
  }

  async loadFacilityData(): Promise<void> {
    try {
      const res: any = await firstValueFrom(
        this.dataService.Get_User_Facility_List_Data(),
      );
      this.facilityListDataSource = res?.data ?? [];

      if (this.facilityListDataSource?.length == 1) {
        // Auto-select first facility
        this.selectedFacility = [
          this.facilityListDataSource[0].FacilityLicense,
        ];
      }
    } catch (error) {
      console.error('Error fetching facility data:', error);
    }
  }

  isApplyButtonDisabled(): boolean {
    return !(this.selectedFacility && this.fromDate && this.toDate);
  }

  // ================= load grid data by using filter values ==============
  onApplyFilter() {
    const formatDate = (date: Date | null) =>
      date ? this.datePipe.transform(date, 'yyyy-MM-dd') : null;

    const payload = {
      FacilityID: Array.isArray(this.selectedFacility)
        ? this.selectedFacility.join(',')
        : '',
      DateFrom: formatDate(this.fromDate),
      DateTo: formatDate(this.toDate),
    };

    this.loadingVisible = true;

    this.operationService.getClinicalData(payload).subscribe({
      next: (res: any) => {
        if (res?.flag === '1') {
          this.xmlGridData = res.data ?? [];

          // Clear previous selections/focus
          this.selectedRowKeys = [];
          this.selectedRowsData = [];
          this.focusedRowKey = null;
        } else {
          this.xmlGridData = [];
        }
      },
      error: (err) => {
        console.error('Error loading data:', err);
        this.xmlGridData = [];
      },
      complete: () => {
        this.loadingVisible = false;
      },
    });
  }

  // ================ Called when a file is selected
  async onFileSelected(event: any, fileInput: HTMLInputElement): Promise<void> {
    // Facility validation
    if (
      !this.selectedFacility ||
      (Array.isArray(this.selectedFacility) &&
        this.selectedFacility.length === 0)
    ) {
      this.notificationService.showNotification(
        'Please select a Facility before importing XML files.',
        'error',
      );

      fileInput.value = '';
      return;
    }

    this.hasError = false;
    this.importResults = [];
    this.isExcelLoading = true;

    const files = event.target.files || [];
    this.totalFiles = files.length;
    this.uploadedCount = 0;
    this.successCount = 0;
    this.alreadyImportedCount = 0;
    this.failCount = 0;

    this.selectedXmlFile = [];

    const processXmlFile = (file: File): Promise<void> => {
      return new Promise((resolve) => {
        const reader = new FileReader();

        reader.onload = () => {
          const base64String = (reader.result as string).split(',')[1];

          const filePayload: any = {
            facilityID: this.selectedFacility.join(','),
            fileName: file.name,
            fileData: base64String,
            userID: 1,
          };

          this.selectedXmlFile.push(filePayload);
          this.isResponsePopupOpened = true;

          this.masterService.ImportClinicalData(filePayload).subscribe({
            next: (res: any) => {
              this.uploadedCount++;

              if (res.message === 'Success') {
                this.successCount++;
              } else if (res.message === 'File already imported.') {
                this.alreadyImportedCount++;
              } else {
                this.failCount++;
              }

              if (Array.isArray(res.data)) {
                this.importResults.push(...res.data);
              }
            },
            error: (err: any) => {
              console.error('Import error:', err);
              this.failCount++;
            },
            complete: () => resolve(),
          });
        };

        reader.readAsDataURL(file);
      });
    };

    // Process XML files only
    for (const file of files) {
      const fileName = file.name.toLowerCase();

      if (!fileName.endsWith('.xml')) {
        notify(
          {
            message: `Only XML files are allowed. (${file.name})`,
            position: { at: 'top right', my: 'top right' },
          },
          'error',
        );
        continue;
      }

      await processXmlFile(file);
    }

    if (this.successCount > 0) {
      const today = new Date();

      this.fromDate = today;
      this.toDate = today;

      this.onApplyFilter();
    }

    this.isExcelLoading = false;
    fileInput.value = '';
  }

  get progressValue() {
    return this.totalFiles > 0
      ? (this.uploadedCount / this.totalFiles) * 100
      : 0;
  }

  format = () => {
    return `Uploaded: ${this.uploadedCount}/${this.totalFiles} | Success: ${this.successCount} | Failed: ${this.failCount} | Already Imported: ${this.alreadyImportedCount}`;
  };

  onStatusCellPrepared(e: any) {
    if (e.rowType === 'data' && e.column.dataField === 'Status') {
      switch ((e.value || '').toLowerCase()) {
        case 'success':
          e.cellElement.style.color = 'green';

          break;
        case 'failed':
          e.cellElement.style.color = 'red';

          break;
        case 'file already imported':
          e.cellElement.style.color = 'orange';
          break;
        default:
          e.cellElement.style.color = 'black';
      }
    }
  }

  // =========== xml datagrid clearing on changing value of upload type =======
  onUploadModeChanged() {
    this.clearSelection();
  }

  // ===== Upload handler =====
  uploadXmlFile() {
    if (!this.selectedRowsData.length) {
      alert('Please select at least one row before uploading!');
      return;
    }

    const userId = sessionStorage.getItem('UserID');
    this.totalCount = this.selectedRowsData.length;
    this.completedCount = 0;
    this.failedCount = 0;
    this.isUploading = true;
    this.isCancelled = false;
    this.currentRequest = undefined;

    this.uploadNextFile(0, userId);
  }

  // ===== Upload logic with cancel check =====
  uploadNextFile(index: number, userId: string | null) {
    if (index >= this.selectedRowsData.length) {
      this.refreshXmlGrid();
      return;
    }

    const row = this.selectedRowsData[index];

    const payload = {
      FacilityID: this.selectedFacility.join('') || '',
      FileName: row.XMLFileName,
      UserID: userId,
      ProcessID: row.ProcessID,
      XMLFileID: row.XMLFileID,
      UploadType: this.uploadMode ? 'test' : 'production',
    };

    console.log(
      `Uploading ${index + 1}/${this.totalCount}: ${row.XMLFileName}`,
    );

    this.currentRequest = this.operationService
      .upload_XML_Data(payload)
      .subscribe({
        next: (res: any) => {
          if (res.flag === '1') {
            this.completedCount++;
          } else {
            this.failedCount++;
            if (this.failedCount === this.totalCount) {
              this.refreshXmlGrid();
              return;
            }

            // Single record special case
            if (this.totalCount === 1) {
              this.finishUpload(res.message || 'Upload failed.', 'error');
              this.refreshXmlGrid();
              return;
            }
          }
          this.handleNext(index, userId);
        },
        error: () => {
          this.failedCount++;
          // Check if all failed
          if (this.failedCount === this.totalCount) {
            this.finishUpload(
              'All uploads failed due to network/server error.',
              'error',
            );
            this.refreshXmlGrid();
            return;
          }

          if (this.totalCount === 1) {
            this.finishUpload('Network or server error.', 'error');
            this.refreshXmlGrid();
            return;
          }

          this.notificationService.showNotification(
            'Network or server error.',
            'error',
          );
          this.handleNext(index, userId);
        },
      });
  }

  // ===== Decide next step =====
  handleNext(index: number, userId: string | null) {
    if (this.isCancelled) {
      this.finishUpload('Upload cancelled by user.', 'warning');
      this.refreshXmlGrid();
      return;
    }
    this.uploadNextFile(index + 1, userId);
  }

  // ===== Cancel button handler =====
  cancelUpload() {
    this.isCancelled = true;
    this.notificationService.showNotification(
      'Upload will stop after current file finishes.',
      'warning',
    );
  }

  // ===== Cleanup / Finish Upload =====
  finishUpload(message: string, type: 'success' | 'error' | 'warning') {
    this.isUploading = false; // ⬅ always stop loading
    this.currentRequest = undefined;

    // Only show a toast if a message exists
    if (message) {
      this.notificationService.showNotification(message, type);
    }
  }

  // ===== XML Refresh =====
  refreshXmlGrid() {
    this.isUploading = false;
    this.clearSelection();
    this.onApplyFilter();
  }

  // ==== Handle details icon click to open XML details popup ====
  onDetailsClick = (e: any) => {
    this.clickedrowData = e.row?.data;

    if (!this.clickedrowData) {
      this.notificationService.showNotification(
        'Unable to retrieve row data.',
        'error',
      );
      return;
    }

    if (!this.clickedrowData.XMLFileID) {
      this.notificationService.showNotification(
        'Invalid XML File ID.',
        'error',
      );
      return;
    }

    const payload = {
      XMLFileID: this.clickedrowData.XMLFileID,
    };

    this.loadingVisible = true;

    this.operationService.get_clicked_Row_Xml_data(payload).subscribe({
      next: (res: any) => {
        if (res?.flag === '1') {
          this.openXmlDetails(res);
        } else {
          this.notificationService.showNotification(
            res?.message || 'Failed to retrieve XML data.',
            'error',
          );
        }
      },
      error: (err) => {
        console.error('XML Details Error:', err);

        this.notificationService.showNotification(
          'Unable to load XML data. Please try again.',
          'error',
        );
      },
      complete: () => {
        this.loadingVisible = false;
      },
    });
  };

  openXmlDetails(row: any) {
    if (!row) {
      this.notificationService.showNotification(
        'Invalid XML response.',
        'error',
      );
      return;
    }

    let xmlData = row.XMLData || row.data?.XMLData;

    if (!xmlData) {
      this.notificationService.showNotification(
        'No XML content available.',
        'warning',
      );
      return;
    }

    const dispositionValue = this.uploadMode ? 'TEST' : 'PRODUCTION';

    xmlData = xmlData.replace(
      /<DispositionFlag>.*<\/DispositionFlag>/,
      `<DispositionFlag>${dispositionValue}</DispositionFlag>`,
    );

    this.selectedXmlRow = row;

    // Format XML before displaying
    this.xmlData = this.formatXml(xmlData);

    this.XmlFileName = this.clickedrowData?.XMLFileName || 'XMLFile.xml';
    this.focusedRowKey = this.clickedrowData?.XMLFileID || null;

    this.showXmlDetailsPopup = true;
  }

  formatXml(xml: string): string {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xml, 'application/xml');

    const serializer = new XMLSerializer();
    const xmlString = serializer.serializeToString(xmlDoc);

    return this.prettyPrintXml(xmlString);
  }

  prettyPrintXml(xml: string): string {
    let formatted = '';
    let pad = 0;

    xml
      .replace(/(>)(<)(\/*)/g, '$1\r\n$2$3')
      .split('\r\n')
      .forEach((node) => {
        let indent = 0;

        if (node.match(/^<\/\w/)) {
          pad--;
        }

        formatted += '  '.repeat(Math.max(pad, 0)) + node + '\r\n';

        if (node.match(/^<[^!?\/][^>]*>$/) && !node.includes('</')) {
          indent = 1;
        }

        pad += indent;
      });

    return formatted;
  }

  // ==== download xml file to system ====
  downloadXmlFile() {
    if (!this.xmlData || !this.XmlFileName) {
      notify({
        message: 'No XML data to download.',
        type: 'warning',
        displayTime: 2000,
        position: 'top right',
      });
      return;
    }

    const blob = new Blob([this.xmlData], { type: 'application/xml' });
    const link = document.createElement('a');
    link.href = window.URL.createObjectURL(blob);
    link.download = this.XmlFileName;
    link.click();
    window.URL.revokeObjectURL(link.href); // cleanup
  }

  onPopupHidden() {
    this.clickedrowData = null;
    this.showXmlDetailsPopup = false;
    this.xmlData = '';
  }

  clearSelection() {
    this.selectedRowKeys = [];
    this.selectedRowsData = [];
    this.focusedRowKey = null;
  }

  //======= dropdown value selection changed =======
  onSelectionChanged(e: any) {
    const validRows = e.selectedRowsData.filter((row: any) =>
      this.isRowSelectable(row),
    );
    this.selectedRowsData = validRows;
    this.selectedRowKeys = validRows.map((row: any) => row.XMLFileName);
    if (this.selectedRowKeys.length !== e.selectedRowKeys.length) {
      e.component.selectRows(this.selectedRowKeys, false);
    }
  }

  //======= Row validation logic =======
  isRowSelectable(row: any): boolean {
    const isTestSuccess =
      row.TestUpload &&
      row.TestUpload.toString().toLowerCase().includes('success');
    const isProdSuccess =
      row.ProductionUpload &&
      row.ProductionUpload.toString().toLowerCase().includes('success');

    if (this.uploadMode === true) {
      return !isTestSuccess;
    } else {
      return isTestSuccess && !isProdSuccess;
    }
  }

  // ========== test upload failed cell click =======
  onCellClick(e: any) {
    // Only for data rows and specific columns
    if (
      e.rowType === 'data' &&
      (e.column?.dataField === 'TestUpload' ||
        e.column?.dataField === 'ProductionUpload')
    ) {
      const cellValue = (e.value || '').toString().toLowerCase();

      if (cellValue.includes('failed')) {
        const rowData = e.data;
        this.loadingVisible = true;

        this.operationService
          .download_XML_error_Data(rowData.XMLFileID)
          .subscribe({
            next: (res: any) => {
              if (res.flag === '1' && res.ErrorReport) {
                const byteCharacters = atob(res.ErrorReport);
                const byteNumbers = new Array(byteCharacters.length);
                for (let i = 0; i < byteCharacters.length; i++) {
                  byteNumbers[i] = byteCharacters.charCodeAt(i);
                }
                const byteArray = new Uint8Array(byteNumbers);
                const blob = new Blob([byteArray], { type: 'application/zip' });

                const link = document.createElement('a');
                link.href = window.URL.createObjectURL(blob);

                // Dynamic filename: Test vs Production
                const prefix =
                  e.column?.dataField === 'TestUpload'
                    ? 'TestErrorReport'
                    : 'ProdErrorReport';

                link.download = `${prefix}_${
                  rowData.XMLFileName || rowData.XMLFileID
                }.zip`;
                link.click();

                notify(
                  { message: 'Error report downloaded', position: 'top right' },
                  'success',
                );
              } else {
                notify(
                  {
                    message: res.message || 'No error report available',
                    position: 'top right',
                  },
                  'warning',
                );
              }
            },
            error: () => {
              notify(
                {
                  message: 'Failed to download error report',
                  position: 'top right',
                },
                'error',
              );
            },
            complete: () => {
              this.loadingVisible = false;
            },
          });
      }
    }
  }

  //========= Page refreshing =========
  refresh = () => {
    this.dataGrid.instance.refresh();
  };

  toggleFilterRow = () => {
    this.isFilterRowVisible = !this.isFilterRowVisible;
  };

  //========================Export data ==========================
  onExporting(event: any) {
    const fileName = 'clinical_data';
    this.service.exportDataGrid(event, fileName);
  }

  displayFacility = (item: any) => {
    if (!item) return '';
    return `${item.FacilityLicense} - ${item.FacilityName}`;
  };
}
@NgModule({
  imports: [
    CommonModule,
    DxDataGridModule,
    DxButtonModule,
    DxDataGridModule,
    DxDropDownButtonModule,
    DxSelectBoxModule,
    DxTextBoxModule,
    DxLookupModule,
    FormPopupModule,
    ClinicalDataImportFormModule,
    DxPopupModule,
    DxDateBoxModule,
    DxDropDownBoxModule,
    DxLoadPanelModule,
    DxProgressBarModule,
    DxTagBoxModule,
    DxTextAreaModule,
    DxAccordionModule,
  ],
  providers: [],
  exports: [],
  declarations: [ClinicalDataComponent],
})
export class ClinicianMajorModule {}
