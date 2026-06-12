import { CommonModule } from '@angular/common';
import { Component, NgModule, OnInit, ViewChild } from '@angular/core';
import {
  DxDataGridComponent,
  DxDropDownBoxComponent,
  DxButtonModule,
  DxDataGridModule,
  DxDropDownButtonModule,
  DxSelectBoxModule,
  DxTextBoxModule,
  DxLookupModule,
  DxResizableModule,
  DxDropDownBoxModule,
  DxFormModule,
  DxDateBoxModule,
  DxToolbarModule,
  DxTagBoxModule,
  DxLoadPanelModule,
  DxRadioGroupModule,
  DxPopupModule,
  DxTextAreaModule,
} from 'devextreme-angular';
import notify from 'devextreme/ui/notify';
import { FormPopupModule } from 'src/app/components';
import { DataService } from 'src/app/services';
import { MasterReportService } from '../../MASTER PAGES/master-report.service';
import { ReportService } from 'src/app/services/Report-data.service';
import { NotificationService } from 'src/app/services/notification.service';

@Component({
  selector: 'app-facility-download-log',
  templateUrl: './facility-download-log.component.html',
  styleUrl: './facility-download-log.component.scss',
  providers: [ReportService, DataService],
})
export class FacilityDownloadLogComponent implements OnInit {
  @ViewChild(DxDataGridComponent, { static: true })
  dataGrid!: DxDataGridComponent;

  @ViewChild('facilityDropDown', { static: false })
  facilityDropDown!: DxDropDownBoxComponent;

  readonly allowedPageSizes: any = [10, 20, 'all'];
  displayMode: any = 'full';
  showPageSizeSelector = true;

  isFilterRowVisible: boolean = false;

  // ========= init data dropdown datasource and value variables =========

  PostOffice_Value: any;
  Facility_DataSource: any[] = [];
  Facility_Value: any[] = [];

  From_Date_Value: any = null;
  To_Date_Value: any = null;
  ToDateMax: any = null;
  minDate: any = null;
  maxDate: Date = new Date();

  isContentVisible: boolean = true;
  gridDataSource: any;

  loadingVisible: boolean = false;
  isDropDownOpened: boolean = false;

  popupVisible = false;
  headerFields: { label: string; value: string }[] = [];
  formattedXml: string = '';

  selectedmonth: any = '';
  selectedYear: any = null;
  monthDataSource: { name: string; value: any }[];
  years: number[] = [];

  TransactionTypeDataSource: any;
  transactionValue: any;

  fileStatusDataSource: any;
  fileStatus: any;

  constructor(
    private dataservice: DataService,
    private masterService: MasterReportService,
    private service: ReportService,
    private notificationService: NotificationService,
  ) {
    this.minDate = new Date(2025, 0, 1);
    this.maxDate = new Date();
    //============Year field dataSource===============
    const currentYear: any = new Date().getFullYear();
    for (let year = currentYear; year >= 2025; year--) {
      this.years.push(year);
    }
    //=============month field datasource============
    this.monthDataSource = this.service
      .getMonths()
      .filter((x: any) => x.value !== '' && x.name !== 'All');
  }

  //=== on init functions of the page ========
  ngOnInit() {
    this.fetch_initData_Values();
    this.get_DropDown_Data();
  }

  // ========== fetch initial dropdown fields datasource ===========
  fetch_initData_Values() {
    const userID = sessionStorage.getItem('UserID');
    this.masterService
      .Get_User_Facility_List_Data(userID)
      .subscribe((response: any) => {
        this.loadingVisible = true;
        if (response.flag == '1') {
          this.Facility_DataSource = response.data;
          this.loadingVisible = false;
          if (this.Facility_DataSource?.length == 1) {
            this.Facility_Value = [this.Facility_DataSource[0].ID];
          }
        }
      });
  }

  get_DropDown_Data() {
    this.masterService
      .Get_GropDown('TRANSACTION_TYPE')
      .subscribe((response: any) => {
        this.TransactionTypeDataSource = response;
      });

    this.masterService
      .Get_GropDown('FILE_STATUS')
      .subscribe((response: any) => {
        this.fileStatusDataSource = response;
      });
  }

  // ============ facility dropdown selection change event =============
  onFacilitySelected(e: any): void {
    if (e.selectedRowsData.length) {
      const selectedFacility = e.selectedRowsData[0];
      this.Facility_Value = [selectedFacility.ID];
      this.PostOffice_Value = selectedFacility.PostOfficeID;
      this.facilityDropDown.instance.close();
    }
  }

  //================ Year value change ===================
  onYearChanged(e: any): void {
    this.selectedYear = e.value;
    this.selectedmonth = '';

    const currentYear = new Date().getFullYear();
    const today = new Date();

    if (this.selectedYear === currentYear) {
      this.From_Date_Value = new Date(this.selectedYear, 0, 1);
      this.To_Date_Value = today;
    } else {
      this.From_Date_Value = new Date(this.selectedYear, 0, 1);
      this.To_Date_Value = new Date(this.selectedYear, 11, 31);
    }
  }

  //================ Month value change ===================
  onMonthValueChanged(e: any): void {
    this.selectedmonth = e.value ?? '';

    const today = new Date();
    const currentYear = today.getFullYear();

    if (this.selectedmonth === '') {
      if (this.selectedYear === currentYear) {
        this.From_Date_Value = new Date(currentYear, 0, 1);
        this.To_Date_Value = today;
      } else {
        this.From_Date_Value = new Date(this.selectedYear, 0, 1);
        this.To_Date_Value = new Date(this.selectedYear, 11, 31);
      }
    } else {
      this.From_Date_Value = new Date(this.selectedYear, this.selectedmonth, 1);

      this.To_Date_Value = new Date(
        this.selectedYear,
        this.selectedmonth + 1,
        0,
      );
    }
  }

  // =========== From Date Change Event =================
  onFromDateChanged(e: any): void {
    this.From_Date_Value = e.value;

    const maxDate = new Date(e.value);
    maxDate.setDate(maxDate.getDate() + 90);

    this.ToDateMax = maxDate > new Date() ? new Date() : maxDate;
  }

  // ====== to date value change event =======
  onToDateChanged(e: any): void {
    const fromDate = new Date(this.From_Date_Value);
    const toDate = new Date(e.value);

    const diffDays = Math.ceil(
      (toDate.getTime() - fromDate.getTime()) / (1000 * 60 * 60 * 24),
    );

    if (diffDays > 90) {
      this.notificationService.showNotification(
        'Date range cannot exceed 90 days.',
        'warning',
      );

      const maxDate = new Date(fromDate);
      maxDate.setDate(maxDate.getDate() + 90);

      this.To_Date_Value = maxDate;
      return;
    }

    this.To_Date_Value = e.value;
  }

  validateDateRange(): void {
    if (!this.From_Date_Value || !this.To_Date_Value) {
      return;
    }

    const fromDate = new Date(this.From_Date_Value);
    const maxToDate = new Date(fromDate);
    maxToDate.setDate(maxToDate.getDate() + 90);

    // Restrict ToDate picker max value
    this.ToDateMax = maxToDate > new Date() ? new Date() : maxToDate;

    if (new Date(this.To_Date_Value) > maxToDate) {
      this.To_Date_Value = maxToDate;
    }
  }

  //================Show and Hide Search parameters==========
  toggleContent() {
    this.isContentVisible = !this.isContentVisible;
  }

  // ==============show or hide filter row of the datagrid =============
  toggleFilterRow = () => {
    this.isFilterRowVisible = !this.isFilterRowVisible;
  };

  //============ refresh datagrid ==============
  refresh = () => {
    this.dataGrid.instance.refresh();
  };

  // ======== format date =============
  formatDate = (date: any): string => {
    if (!date) return '';
    const d = new Date(date);
    const year = d.getFullYear();
    const month = `0${d.getMonth() + 1}`.slice(-2);
    const day = `0${d.getDate()}`.slice(-2);
    return `${year}-${month}-${day}`;
  };

  isApplyButtonDisabled(): boolean {
    return !(
      this.Facility_Value &&
      this.From_Date_Value &&
      this.transactionValue &&
      this.fileStatus &&
      this.To_Date_Value
    );
  }

  // ================= fetch datagrid data ===============
  get_Datagrid_DataSource() {
    const fromDate = new Date(this.From_Date_Value);
    const toDate = new Date(this.To_Date_Value);

    const diffMonths =
      (toDate.getFullYear() - fromDate.getFullYear()) * 12 +
      (toDate.getMonth() - fromDate.getMonth());

    // If more than 3 months or invalid range
    if (diffMonths > 3 || fromDate > toDate) {
      notify('Please select a date range within 3 months.', 'warning', 3000);
      return;
    }

    this.loadingVisible = true;

    const inputData = {
      FacilityID: this.Facility_Value.join(','),
      DateFrom: this.formatDate(this.From_Date_Value),
      DateTo: this.formatDate(this.To_Date_Value),
      TransactionType: this.transactionValue,
      FileStatus: this.fileStatus,
    };

    this.dataservice.get_facility_download_Data(inputData).subscribe({
      next: (res: any) => {
        if (res.flag === '1') {
          this.gridDataSource = res.data;
        } else {
          this.gridDataSource = [];
          notify(res.message || 'No data available.', 'error', 3000);
        }
        this.loadingVisible = false;
      },
      error: (err) => {
        this.gridDataSource = [];
        this.loadingVisible = false;
        notify('An error occurred while loading data.', 'error', 3000);
        console.error('API error:', err);
      },
    });
  }

  // =========== detail data fetching ================
  openDetailsPopup = (event: any) => {
    this.loadingVisible = true;

    const selectedRowData = event.row.data;
    const inputFiledata = {
      FacilityID: this.Facility_Value.join(','),
      FileID: selectedRowData.FileID,
    };

    this.dataservice.get_facility_File_Data(inputFiledata).subscribe({
      next: (res: any) => {
        if (res.flag === '1') {
          const xmlString = res.XMLData;
          this.formattedXml = xmlString;

          this.headerFields = [];

          // Add File Name as the first item
          this.headerFields.push({
            label: 'File Name',
            value: selectedRowData.FileName || '',
          });

          // Extract Header fields from XML
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(xmlString, 'text/xml');
          const headerNode = xmlDoc.getElementsByTagName('Header')[0];

          if (headerNode) {
            Array.from(headerNode.children).forEach((child) => {
              this.headerFields.push({
                label: child.nodeName,
                value: child.textContent ?? '',
              });
            });
          }

          this.popupVisible = true;
        } else {
          notify('Failed to fetch file data. Please try again.', 'error', 3000);
        }
      },
      error: (err) => {
        console.error('API error:', err);
        notify(
          'API error occurred. Please check your connection.',
          'error',
          3000,
        );
      },
      complete: () => {
        this.loadingVisible = false;
      },
    });
  };

  // ======== download the popup xml content ============
  downloadXmlFile(): void {
    const blob = new Blob([this.formattedXml], { type: 'application/xml' });
    // Get file name from headerFields
    const fileNameEntry = this.headerFields.find(
      (field) => field.label === 'File Name',
    );
    let fileName = fileNameEntry?.value || 'file';
    // Remove existing .xml/.XML if already present
    if (fileName.toLowerCase().endsWith('.xml')) {
      fileName = fileName.slice(0, -4);
    }
    fileName += '.xml';
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
    // Clean up
    URL.revokeObjectURL(link.href);
  }

  displayFacility = (item: any) => {
    if (!item) return '';
    return `${item.FacilityLicense} - ${item.FacilityName}`;
  };
}

@NgModule({
  imports: [
    DxButtonModule,
    DxDataGridModule,
    DxDropDownButtonModule,
    DxSelectBoxModule,
    DxTextBoxModule,
    DxLookupModule,
    DxResizableModule,
    DxDropDownBoxModule,
    FormPopupModule,
    CommonModule,
    DxFormModule,
    DxDateBoxModule,
    DxToolbarModule,
    DxTagBoxModule,
    DxLoadPanelModule,
    DxRadioGroupModule,
    DxPopupModule,
    DxTextAreaModule,
  ],
  providers: [],
  exports: [],
  declarations: [FacilityDownloadLogComponent],
})
export class FacilityDownloadLogModule {}
