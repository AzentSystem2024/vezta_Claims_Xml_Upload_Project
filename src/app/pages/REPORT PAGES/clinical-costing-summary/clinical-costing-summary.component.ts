import { CommonModule } from '@angular/common';
import { Component, NgModule, ViewChild } from '@angular/core';
import {
  DxDataGridComponent,
  DxTreeViewComponent,
  DxLookupComponent,
  DxButtonModule,
  DxDataGridModule,
  DxSelectBoxModule,
  DxTextBoxModule,
  DxDropDownBoxModule,
  DxFormModule,
  DxDateBoxModule,
  DxTreeViewModule,
  DxValidatorModule,
  DxValidationSummaryModule,
  DxLoadPanelModule,
  DxPopupModule,
  DxDropDownBoxComponent,
} from 'devextreme-angular';
import { DataSource } from 'devextreme/common/data';
import notify from 'devextreme/ui/notify';
import { FormPopupModule } from 'src/app/components';
import { ReportService } from 'src/app/services/Report-data.service';
import {
  CptMasterEditFormComponent,
  CptMasterEditFormModule,
} from '../../POP-UP_PAGES/cpt-master-edit-form/cpt-master-edit-form.component';
import { MasterReportService } from '../../MASTER PAGES/master-report.service';
import { OperationReportService } from '../../OPERATION PAGES/operation-report.service';
import {
  ClinicianEditFormComponent,
  ClinicianEditFormModule,
} from '../../POP-UP_PAGES/clinician-edit-form/clinician-edit-form.component';
import { SingleClaimDetailsModule } from '../../REPORT POPUP PAGES/single-claim-details/single-claim-details.component';
import { ReportEngineService } from '../report-engine.service';
import { forkJoin } from 'rxjs';

@Component({
  selector: 'app-clinical-costing-summary',

  templateUrl: './clinical-costing-summary.component.html',
  styleUrl: './clinical-costing-summary.component.scss',
  providers: [ReportService, ReportEngineService],
})
export class ClinicalCostingSummaryComponent {
  @ViewChild(DxDataGridComponent, { static: true })
  dataGrid: DxDataGridComponent;

  @ViewChild('facilityDropDown', { static: false })
  facilityDropDown: DxDropDownBoxComponent;
  @ViewChild('processDropDown', { static: false })
  processDropDown: DxDropDownBoxComponent;

  //=================DataSource for data Grid Table========
  dataGrid_DataSource: DataSource<any>;
  columnsConfig: any;

  //================Variables for Storing DataSource========

  Facility_DataSource: any;
  Facility_Value: any;
  allProcessData: any;
  Process_DataSource: any;
  ProcessID_Value: any;
  //========Variables for Pagination ====================
  readonly allowedPageSizes: any = [10, 20, 'all'];
  displayMode: any = 'full';
  showPageSizeSelector = true;
  showInfo = true;
  showNavButtons = true;
  show_Pagination = true;

  //=====================other variables==================
  isContentVisible: boolean = true;
  currentPathName: any;

  isFilterOpened = false; //filter row enable-desable variable
  isDetailFilterOpened = false;
  GridSource: any;
  isEmptyDatagrid: boolean = true;
  summaryColumnsData: any;
  columndata: any;

  loadingVisible: boolean = false;
  userID: any;
  selectedData: any;
  SummaryDetailsDatagrid:any;
  ViewCostingSummaryPopup:boolean = false;
  Particulars: any;

  constructor(
    private service: ReportService,
    private reportengine: ReportEngineService,
    private masterService: MasterReportService,
    private operationService: OperationReportService
  ) {
    this.get_searchParameters_Dropdown_Values();
  }

  //================Show and Hide Search parameters========
  toggleContent() {
    this.isContentVisible = !this.isContentVisible;
  }

  //============Get search parameters dropdown values=======
  get_searchParameters_Dropdown_Values(): void {
    this.userID = sessionStorage.getItem('UserID');
    if (!this.userID) return;

    // this.loadingVisible = true;

    // Reordered: processRequest first, facilityRequest second
    const processRequest = this.operationService.getClinicalCostingList(
      this.userID
    );
    const facilityRequest = this.masterService.Get_User_Facility_List_Data(
      this.userID
    );

    forkJoin([processRequest, facilityRequest]).subscribe({
      next: ([processRes, facilityRes]: any) => {
        // Process Data
        if (processRes?.flag === '1') {
          this.allProcessData = processRes.data || [];
        }

        // Facility Data
        if (facilityRes?.flag === '1') {
          this.Facility_DataSource = facilityRes.data || [];

          // Auto-select when 1 facility
          if (this.Facility_DataSource.length === 1) {
            const singleFacility = this.Facility_DataSource[0].FacilityLicense;
            this.Facility_Value = [singleFacility];

            // Trigger cascading to process dropdown
            this.onFacilityChanged({ value: singleFacility });
          }
        }
      },
      error: () => {
        console.error('Error loading dropdown values');
      },
      complete: () => {
        // this.loadingVisible = false;
      },
    });
  }

  onFacilityChanged(e: any) {
    const selectedFacility = e.value;
    console.log('Selected Facility :>', selectedFacility);

    this.Facility_Value = selectedFacility;

    // Reset process values
    this.ProcessID_Value = null;
    this.Process_DataSource = []; // force clear before refill

    if (selectedFacility) {
      this.facilityDropDown.instance.close();

      // Filter the new Process list
      const filteredData = this.allProcessData.filter(
        (p: any) => p.FacilityID === selectedFacility
      );

      // Important: Reassign to trigger Angular change detection
      this.Process_DataSource = [...filteredData];

      console.log('Filtered Process_DataSource:', this.Process_DataSource);
    } else {
      this.Process_DataSource = [];
    }

    // 🔧 Optional: Reset the grid selection if open
    if (this.processDropDown) {
      this.processDropDown.instance.option('value', null);
    }
  }

  onFacilityGridSelectionChanged(e: any) {
    const selectedFacility = e.selectedRowKeys[0];
    console.log('Facility selected from grid:', selectedFacility);

    this.Facility_Value = selectedFacility;
    this.facilityDropDown.instance.close();

    // Reset Process selection
    this.ProcessID_Value = null;
    this.Process_DataSource = [];

    if (selectedFacility) {
      // Filter processes based on FacilityID (matches FacilityLicense)
      const filteredData = this.allProcessData.filter(
        (p: any) => p.FacilityID === selectedFacility
      );

      // Important: create a new array reference for change detection
      this.Process_DataSource = [...filteredData];
      console.log('Filtered Process Data:', this.Process_DataSource);

      // Optional: close and reset process drop-down if open
      if (this.processDropDown) {
        this.processDropDown.instance.option('value', null);
        this.processDropDown.instance.close();
      }
    }
  }


  onProcessChanged(e: any) {
    const selectedProcess = e.value;
    this.ProcessID_Value = selectedProcess;
    if (selectedProcess) {
      this.processDropDown.instance.close();
    }
  }

  // ========= Fetch DataSource for the Main DataGrid Table =========
  async get_Datagrid_DataSource() {
    const ProcessID = Array.isArray(this.ProcessID_Value)
      ? Number(this.ProcessID_Value[0])
      : Number(this.ProcessID_Value);

    console.log('Process ID:', ProcessID);

    const formData = { ProcessID };

    // this.isContentVisible = false;
    this.dataGrid.instance.beginCustomLoading('Loading...');

    try {
      const response: any = await this.service
        .fetch_Clinical_Costing_Summary_Data(formData)
        .toPromise();

      if (response?.flag === '1') {
        this.isEmptyDatagrid = response.data?.length === 0;
        this.dataGrid_DataSource = response.data || [];

        this.dataGrid.instance.endCustomLoading();
        // this.isContentVisible = false;
      } else {
        this.dataGrid.instance.endCustomLoading();
        // this.isContentVisible = true;

        notify(
          {
            message: response?.message || 'Something went wrong!',
            position: { at: 'top right', my: 'top right' },
          },
          'error'
        );
      }
    } catch (error) {
      this.dataGrid.instance.endCustomLoading();
      // this.isContentVisible = true;

      notify(
        {
          message:
            'An error occurred while fetching the data. Please try again later.',
          position: { at: 'top right', my: 'top right' },
          displayTime: 3000,
        },
        'error'
      );
    }
  }


  viewDetails = (e: any) => {
  const selectedData = e.row.key;
  this.selectedData = selectedData;
  this.Particulars = selectedData.Particulars;

  const payload = {
    ID: selectedData.SlNo,
    ProcessID: Number(this.ProcessID_Value)
  };

  // Initialize the DataSource for the popup grid
  this.SummaryDetailsDatagrid = new DataSource<any>({
    load: () =>
      new Promise((resolve, reject) => {
        this.service.fetch_Clinical_Costing_Summary_Data_Details(payload).subscribe({
          next: (res: any) => resolve(res.data || []),
          error: ({ message }) => reject(message),
        });
      }),
  });

  // Open the popup
  this.ViewCostingSummaryPopup = true;
};



  //============ Show Filter Row ============
  filterClick = () => {
    if (this.dataGrid_DataSource) {
      this.isFilterOpened = !this.isFilterOpened;
    }
  };

  DetailsfilterClick = () => {
    if (this.SummaryDetailsDatagrid) {
      this.isDetailFilterOpened = !this.isDetailFilterOpened;
    }
  };

  //============Show Filter Row==========================
  SummaryClick = () => {
    const reportGridElement = document.querySelector('.reportGrid');
    if (reportGridElement) {
      reportGridElement.classList.toggle('reportGridFooter');
    }
  };

  //=====================Search on Each Column===========
  applyFilter() {
    this.GridSource.filter();
  }

  //====================Find the column location from the datagrid================
  findColumnLocation = (e: any) => {
    const columnName = e.itemData;
    if (columnName != '' && columnName != null) {
      this.reportengine.makeColumnVisible(this.dataGrid, columnName);
    }
  };

  //=============DataGrid Refreshing=====================
  refresh = () => {
    this.dataGrid.instance.refresh();
  };

  //============= Exporting Function ==============
  onExporting(event: any) {
    const fileName = 'Clinical Costing Summary';
    this.service.exportDataGrid(event, fileName);
  }

  onExportingDetails(event: any) {
    const fileName = 'Clinical Costing Summary Details';
    this.service.exportDataGrid(event, fileName);
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
    DxSelectBoxModule,
    DxTextBoxModule,
    DxDropDownBoxModule,
    CommonModule,
    DxFormModule,
    DxDateBoxModule,
    DxTreeViewModule,
    DxValidatorModule,
    DxValidationSummaryModule,
    DxLoadPanelModule,
    FormPopupModule,
    DxPopupModule,
    CptMasterEditFormModule,
    ClinicianEditFormModule,
    SingleClaimDetailsModule,
  ],
  providers: [],
  exports: [],
  declarations: [ClinicalCostingSummaryComponent],
})
export class ClinicalCostingSummaryModule {}
