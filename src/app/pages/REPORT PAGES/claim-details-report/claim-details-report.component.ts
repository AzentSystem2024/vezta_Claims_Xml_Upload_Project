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
} from 'devextreme-angular';
import { CustomStore, DataSource } from 'devextreme/common/data';
import notify from 'devextreme/ui/notify';
import { ReportService } from 'src/app/services/Report-data.service';
import { MasterReportService } from '../../MASTER PAGES/master-report.service';
import { ReportEngineService } from '../report-engine.service';
import { Workbook } from 'exceljs';
import * as FileSaver from 'file-saver';
import {
  ClinicianEditFormModule,
  ClinicianEditFormComponent,
} from '../../POP-UP_PAGES/clinician-edit-form/clinician-edit-form.component';
import {
  CptMasterEditFormComponent,
  CptMasterEditFormModule,
} from '../../POP-UP_PAGES/cpt-master-edit-form/cpt-master-edit-form.component';
import { FormPopupModule } from 'src/app/components';

import {
  SingleClaimDetailsComponent,
  SingleClaimDetailsModule,
} from '../../REPORT POPUP PAGES/single-claim-details/single-claim-details.component';
import { OperationReportService } from '../../OPERATION PAGES/operation-report.service';
@Component({
  selector: 'app-claim-details-report',
  templateUrl: './claim-details-report.component.html',
  styleUrl: './claim-details-report.component.scss',
  providers: [ReportService, ReportEngineService],
})
export class ClaimDetailsReportComponent {
  @ViewChild(DxDataGridComponent, { static: true })
  dataGrid: DxDataGridComponent;

  @ViewChild(DxTreeViewComponent, { static: false })
  treeView: DxTreeViewComponent;

  @ViewChild(CptMasterEditFormComponent, { static: false })
  CptEditFormComponent: CptMasterEditFormComponent;

  @ViewChild(ClinicianEditFormComponent, { static: false })
  clinicianEditComponent: ClinicianEditFormComponent;

  @ViewChild('lookup', { static: false }) lookup: DxLookupComponent;

  //=================DataSource for data Grid Table========
  dataGrid_DataSource: DataSource<any>;

  columnsConfig: any; //==============Column data storing variable

  //================Variables for Storing DataSource========

  Facility_DataSource: any;

  monthDataSource: { name: string; value: any }[];
  years: number[] = [];

  //================Variables for Storing selected Parameters========

  Facility_Value: any;

  From_Date_Value: any = new Date();
  To_Date_Value: any = new Date();

  selectedmonth: any = '';
  selectedYear: number | null = null;

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

  minDate: Date;
  maxDate: Date;

  isFilterOpened = false; //filter row enable-desable variable
  GridSource: any;
  isEmptyDatagrid: boolean = true;
  summaryColumnsData: any;
  columndata: any;

  loadingVisible: boolean = false;

  exportingVisible:boolean = false;

  //============Custom close button for drilldown popup============

  userID: any;
  DetailData: any;
  DetailDataCaption: any;
  FilteredDetailData: any;
  DetailsColumns: any;
  DetailsummaryColumns: any;

  originalHeaderData: any[] = [];
  originalDetailData: any[] = [];
  searchOnDataSource = [
    { id: 'EncounterStartDate', name: 'Encounter Start Date' },
    { id: 'EncounterEndDate', name: 'Encounter End Date' },
    { id: 'TransactionDate', name: 'Transaction Date' },
  ];
  selectedSearchOn: any = 'EncounterEndDate';

  detailCache: { [claimUID: string]: any[] } = {};
  selectedCptCodeData: any;
  selectedClinicianData: any;
  isCptEditFormPopupOpened: boolean = false;
  isEditClinicianPopupOpened: boolean = false;

  toolbarEditItems: any = [
    {
      widget: 'dxButton',
      options: {
        text: 'Cancel',
        stylingMode: 'outlined',
        type: 'normal',
        onClick: () => {
          this.isEditClinicianPopupOpened = false;
        },
      },
      toolbar: 'bottom',
      location: 'after',
    },
    {
      widget: 'dxButton',
      options: {
        text: 'Save',
        type: 'default',
        stylingMode: 'contained',
        onClick: () => this.onClickUpdateNewClinician(),
      },
      toolbar: 'bottom',
      location: 'after',
    },
  ];
  isSingleClaimDetailsVisible: boolean = false;
  clickedCellRowData: any;

  constructor(
    private service: ReportService,
    private reportengine: ReportEngineService,
    private masterService: MasterReportService,
    private operationService: OperationReportService
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
    this.get_searchParameters_Dropdown_Values();
  }

  //================Show and Hide Search parameters========
  toggleContent() {
    this.isContentVisible = !this.isContentVisible;
  }

  //============Get search parameters dropdown values=======
  get_searchParameters_Dropdown_Values() {
    this.userID = sessionStorage.getItem('UserID');
    this.masterService
      .Get_User_Facility_List_Data(this.userID)
      .subscribe((response: any) => {
        if (response.flag == '1') {
          this.Facility_DataSource = response.data;
          this.loadingVisible = false;

          if (this.Facility_DataSource?.length == 1) {
            // Auto-select first facility
            this.Facility_Value = [this.Facility_DataSource[0].FacilityLicense];
          }
        }
      });
  }
  // ========= Fetch DataSource for the Main DataGrid Table =========
  async get_Datagrid_DataSource() {
    const formData = {
      FacilityID: this.Facility_Value.join(','),
      SearchOn: this.selectedSearchOn,
      DateFrom: this.reportengine.formatDate(this.From_Date_Value),
      DateTo: this.reportengine.formatDate(this.To_Date_Value),
      ClaimUID: 0,
    };

    this.isContentVisible = false;
    this.dataGrid.instance.beginCustomLoading('Loading...');

    try {
      const response: any = await this.service
        .fetch_Claim_Details_Data(formData)
        .toPromise();

      if (response.flag === '1') {
        this.isEmptyDatagrid = false;

        this.columndata = response.header.ReportColumns;

        const userLocale = navigator.language || 'en-US';

        this.summaryColumnsData = this.generateSummaryColumns(
          response.header.ReportColumns
        );
        this.columnsConfig = this.generateColumnsConfig(
          response.header.ReportColumns,
          userLocale
        );

        this.originalHeaderData = response.header.ReportData;
        this.detailCache = {}; // 🔁 initialize detail data cache

        this.dataGrid_DataSource = new DataSource<any>({
          load: () => Promise.resolve(this.originalHeaderData),
        });

        this.dataGrid.instance.endCustomLoading();
        this.isContentVisible = false;
      } else {
        this.dataGrid.instance.endCustomLoading();
        this.isContentVisible = false;
        notify(
          {
            message: `${response.message}`,
            position: { at: 'top right', my: 'top right' },
          },
          'error'
        );
      }
    } catch (error) {
      this.dataGrid.instance.endCustomLoading();
      this.isContentVisible = true;
      notify(
        {
          message: `An error occurred while fetching the data. Please try again later.`,
          position: { at: 'top right', my: 'top right' },
          displayTime: 3000,
        },
        'error'
      );
    }
  }

  // ========= Main DataGrid Row Expanding Function =========
  async headerDataRowExpanding(e: any) {

    this.dataGrid.instance.collapseAll(-1);

    const claimUid = e.key.ClaimUID || e.data.ClaimUID;
    const userLocale = navigator.language || 'en-US';

    this.FilteredDetailData = [];
    //  If data already exists in cache, use it
    if (this.detailCache?.[claimUid]) {
      this.FilteredDetailData = this.detailCache[claimUid];
    } else {
      // 📡 Fetch detail data only for the expanded row
      const formData = {
        FacilityID: this.Facility_Value.join(','),
        SearchOn: this.selectedSearchOn,
        DateFrom: this.reportengine.formatDate(this.From_Date_Value),
        DateTo: this.reportengine.formatDate(this.To_Date_Value),
        ClaimUID: claimUid,
      };

      this.isContentVisible = false;
      this.dataGrid.instance.beginCustomLoading('Loading...');

      try {
        const response: any = await this.service
          .fetch_Claim_Details_Data(formData)
          .toPromise();

        if (response.flag === '1') {
          const detailData = response.detail.ReportData;

          if (!this.DetailsColumns?.length) {
            this.DetailDataCaption = response.detail.ReportID;
            this.DetailsummaryColumns = this.generateSummaryColumns(
              response.detail.ReportColumns
            );
            this.DetailsColumns = this.generateColumnsConfig(
              response.detail.ReportColumns,
              userLocale
            );
          }

          //  Cache and assign the data
          this.detailCache[claimUid] = detailData;
          this.FilteredDetailData = detailData;
          this.originalDetailData = detailData;
        } else {
          notify(
            {
              message: `Detail loading failed: ${response.message}`,
              position: { at: 'top right', my: 'top right' },
            },
            'error'
          );
        }
      } catch (error) {
        notify(
          {
            message: `An error occurred while fetching detail data.`,
            position: { at: 'top right', my: 'top right' },
            displayTime: 3000,
          },
          'error'
        );
      } finally {
        this.dataGrid.instance.endCustomLoading();
      }
    }

    // // ✅ Only show expanded row (as per requirement)
    // const expandedRow = this.originalHeaderData.find(
    //   (row) => row.ClaimUID === claimUid
    // );
    // this.dataGrid_DataSource = new DataSource<any>({
    //   load: () => Promise.resolve([expandedRow]),
    // });

    // requestAnimationFrame(() => {
    //   const rowIndex = this.dataGrid.instance.getRowIndexByKey(e.key);
    //   if (rowIndex >= 0) {
    //     const scrollable = (this.dataGrid.instance as any).getScrollable();
    //     const rowElements: any = this.dataGrid.instance.getRowElement(rowIndex);
    //     if (scrollable && rowElements?.length > 0) {
    //       const topOffset = rowElements[0].offsetTop;
    //       scrollable.scrollTo({ top: topOffset });
    //     }
    //   }
    // });
  }

  // ========= Main DataGrid Row Collapsing Function =========
  onRowCollapsing(e: any) {
    this.dataGrid_DataSource = new DataSource<any>({
      load: () => Promise.resolve(this.originalHeaderData),
    });
  }

  // =========== Datagrid column formating =============
  generateColumnsConfig(reportColumns, userLocale) {
    return reportColumns.map((column) => {
      let columnFormat;

      if (column.Type === 'DateTime') {
        columnFormat = {
          type: 'date',
          formatter: (date) =>
            new Intl.DateTimeFormat(userLocale, {
              year: 'numeric',
              month: 'short',
              day: '2-digit',
            }).format(new Date(date)),
        };
      }

      if (column.Type === 'Decimal') {
        columnFormat = {
          formatter: (value: number) => {
            if (value == null) return '';

            return new Intl.NumberFormat('en-US', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            }).format(value);
          },
        };
      }

      if (column.Type === 'percentage') {
        columnFormat = {
          type: 'percent',
          precision: 2,
          formatter: (value) =>
            new Intl.NumberFormat(userLocale, {
              style: 'percent',
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            }).format(value / 100),
        };
      }

      return {
        dataField: column.Name,
        caption: column.Title,
        visible: column.Visibility,
        type: column.Type,
        format: columnFormat,
      };
    });
  }

  // =========== Summary column formatting =============
  generateSummaryColumns(reportColumns: any) {
    const decimalColumns = reportColumns.filter(
      (col) => col.Type === 'Decimal' && col.Summary
    );

    const intColumns = reportColumns.filter(
      (col) => col.Type === 'Int32' && col.Summary
    );

    return {
      totalItems: [
        ...decimalColumns.map((col) =>
          this.createSummaryItem(col, false, 'sum', 'decimal')
        ),
        ...intColumns.map((col) =>
          this.createSummaryItem(col, false, 'sum', 'count')
        ),
      ],
      groupItems: [
        ...decimalColumns.map((col) =>
          this.createSummaryItem(col, true, 'sum', 'decimal')
        ),
        ...intColumns.map((col) =>
          this.createSummaryItem(col, true, 'sum', 'count')
        ),
      ],
    };
  }

  // =========== Custom summary item with same smart decimal logic ==========
  createSummaryItem(col, isGroupItem = false, summaryType = 'sum', formatType) {
    return {
      column: col.Name,
      summaryType: summaryType,
      displayFormat: formatType === 'count' ? 'Count: {0}' : '{0}',
      valueFormat:
        formatType === 'decimal'
          ? {
              formatter: (value) => {
                if (value == null) return '';

                const fixed = value.toFixed(3);
                const decimals = fixed.split('.')[1] || '000';

                // Decide 2 or 3 decimal places
                const decimalPlaces = decimals[2] === '0' ? 2 : 3;

                // Format number in US style, no currency, no grouping
                return new Intl.NumberFormat('en-US', {
                  minimumFractionDigits: decimalPlaces,
                  maximumFractionDigits: decimalPlaces,
                }).format(value);
              },
            }
          : {
              formatter: (value) =>
                value != null
                  ? new Intl.NumberFormat('en-US', {
                      useGrouping: false,
                    }).format(Number(value))
                  : '',
            },
      alignByColumn: isGroupItem,
      showInGroupFooter: isGroupItem,
    };
  }

  // ======= cpt code and ordering clinician edit function ============
  onCellClick(e: any) {
    if (!e.column || e.rowType !== 'data') {
      return;
    }

    if (e.rowType === 'group') return;

    const dataField = e.column.dataField;

    // --- Helper to avoid repeated notify options---
    const showError = (message: string) => {
      notify(message, 'error', 3000);
    };

    // ===== Check for Claim Number click =====
    if (dataField === 'ClaimNumber') {
      this.clickedCellRowData = e.data;
      this.isSingleClaimDetailsVisible = true;
      return;
    }

    if (dataField === 'CPTCode') {
      const code = e.data?.CPTCode;

      if (!code) {
        showError('CPT Code is empty');
        return;
      }

      this.operationService
        .fetch_selected_CptCode_Data(code)
        .subscribe((res: any) => {
          if (res.flag === '1' && res.data?.[0]) {
            this.selectedCptCodeData = res.data[0];
            this.isCptEditFormPopupOpened = true;
          } else {
            showError('No CPT Code data found');
          }
        });
    }

    if (dataField === 'OrderingClinician') {
      const clinicianId = e.data?.OrderingClinician;

      if (!clinicianId) {
        showError('Ordering Clinician is empty');
        return;
      }

      this.operationService
        .fetch_selected_orderingClinician_Data(clinicianId)
        .subscribe((res: any) => {
          if (res.flag === '1' && res.data?.[0]) {
            this.selectedClinicianData = res.data[0];
            this.isEditClinicianPopupOpened = true;
          } else {
            showError('No Ordering Clinician data found');
          }
        });
    }

    if (dataField === 'Clinician') {
      const clinicianId = e.data?.Clinician;

      if (!clinicianId) {
        showError('Clinician is empty');
        return;
      }
      this.operationService
        .fetch_selected_orderingClinician_Data(clinicianId)
        .subscribe((res: any) => {
          if (res.flag === '1' && res.data?.[0]) {
            this.selectedClinicianData = res.data[0];
            this.isEditClinicianPopupOpened = true;
          } else {
            showError('No Clinician data found');
          }
        });
    }
  }

  // =========== update Cpt data ===========
  onClickUpdateNewCptType = () => {
    const cptData = this.CptEditFormComponent.getUpdateCptMasterData();

    this.CptEditFormComponent.newCptMasterData.selectedLedgerID =
  this.CptEditFormComponent.ledgerMode === 1
    ? this.CptEditFormComponent.selectedLedgerIds.join(',')
    : '';

    const {
      ID,
      CPTTypeID,
      CPTCode,
      CPTName,
      Description,
      CPTGroup,
      DepartmentID,
      CPTDepartmentID,
      CostDepartmentID,
      CostDriveID,
      FixedQuantity,
      IsDifferentCPTDepartment,
      IsDifferentLedger,     
      selectedLedgerID,    
      CPTEncounterDepartments,
      data,
    } = this.CptEditFormComponent.getUpdateCptMasterData();

    this.masterService
      .update_CptMaster_data(
        ID,
        CPTTypeID,
        CPTCode,
        CPTName,
        Description,
        CPTGroup,
        DepartmentID,
        CPTDepartmentID,
        CostDepartmentID,
        CostDriveID,
        FixedQuantity,
        IsDifferentCPTDepartment,
        IsDifferentLedger,     
        selectedLedgerID,    
        CPTEncounterDepartments,
        data
      )
      .subscribe((response: any) => {
        if (response) {
          this.dataGrid.instance.refresh();
          notify(
            {
              message: `Cpt Master Updated Successfully`,
              position: { at: 'top right', my: 'top right' },
            },
            'success'
          );
        } else {
          notify(
            {
              message: `Your Data Not Updated`,
              position: { at: 'top right', my: 'top right' },
            },
            'error'
          );
        }
      });
  };

  // =========== update oredering clinician =========
  onClickUpdateNewClinician = () => {
    const data = this.clinicianEditComponent.getnewClinicianData();
    console.log(data, 'data');
    const {
      ID,
      ClinicianLicense,
      ClinicianName,
      ClinicianShortName,
      SpecialityID,
      MajorID,
      ProfessionID,
      CategoryID,
      Gender,
      DepartmentID,
    } = this.clinicianEditComponent.getnewClinicianData();

    this.masterService
      .update_Clinician_data(
        ID,
        ClinicianLicense,
        ClinicianName,
        ClinicianShortName,
        SpecialityID,
        MajorID,
        ProfessionID,
        CategoryID,
        Gender,
        DepartmentID
      )
      .subscribe((response: any) => {
        if (response) {
          notify(
            {
              message: `Clinician updated Successfully`,
              position: { at: 'top right', my: 'top right' },
            },
            'success'
          );
          this.isEditClinicianPopupOpened = false;
          this.dataGrid.instance.refresh();
        } else {
          notify(
            {
              message: `Your Data Not Saved`,
              position: { at: 'top right', my: 'top right' },
            },
            'error'
          );
        }
      });
  };

  //============ Show Filter Row ============
  filterClick = () => {
    if (this.dataGrid_DataSource) {
      this.isFilterOpened = !this.isFilterOpened;
    }
  };

  //============Show Filter Row==========================
  SummaryClick = () => {
    const reportGridElement = document.querySelector('.reportGrid');
    if (reportGridElement) {
      reportGridElement.classList.toggle('reportGridFooter');
    }
  };
  //================ Year value change ===================
  onYearChanged(e: any): void {
    this.selectedYear = e.value;
    this.selectedmonth = '';
    const currentYear = new Date().getFullYear();
    const today = new Date();
    if (this.selectedYear === currentYear) {
      // Set from date to the start of the year and to date to today
      this.From_Date_Value = new Date(this.selectedYear, 0, 1); // January 1 of the current year
      this.To_Date_Value = today; // Today's date
    } else {
      this.From_Date_Value = new Date(this.selectedYear, 0, 1); // January 1
      this.To_Date_Value = new Date(this.selectedYear, 11, 31); // December 31
    }
  }

  //================Month value change ===================
  onMonthValueChanged(e: any) {
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
        0
      );
    }
  }

  //============Hide drop down after Value Selected======
  onDropdownValueChanged() {
    const lookupInstance = this.lookup.instance;
    if (lookupInstance) {
      lookupInstance.close();
      lookupInstance.option('searchValue', '');
    }
  }

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
    const fileName = 'Cliam Details';
    this.service.exportDataGrid(event, fileName);
  }

  exportServerExcel = () => {
    
      const payload = {
        reportId: 'rpt-claim-details',
        facilityId: this.Facility_Value.join(','),
        periodFrom: this.reportengine.formatDate(this.From_Date_Value),
        periodTo: this.reportengine.formatDate(this.To_Date_Value),
        SearchOn : this.selectedSearchOn
      };
    
      this.exportingVisible = true;
    
      this.masterService.exportReport(payload).subscribe({
        next: (blob: Blob) => {
          const fileName =
            `ClaimDetails_${new Date().toISOString().slice(0, 19)}.xlsx`;
    
          FileSaver.saveAs(blob, fileName);
          this.exportingVisible = false;
        },
        error: () => {
          this.exportingVisible = false;
          notify('Export failed. Please try again.', 'error', 3000);
        }
      });
    };

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
  declarations: [ClaimDetailsReportComponent],
})
export class ClaimDetailsReportModule {}
