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
import { DataSource } from 'devextreme/common/data';
import notify from 'devextreme/ui/notify';
import { Workbook } from 'exceljs';
import * as FileSaver from 'file-saver';
import { ReportService } from 'src/app/services/Report-data.service';
import { MasterReportService } from '../../MASTER PAGES/master-report.service';
import { ReportEngineService } from '../report-engine.service';
import { OperationReportService } from '../../OPERATION PAGES/operation-report.service';
import {
  ClinicianEditFormModule,
  ClinicianEditFormComponent,
} from '../../POP-UP_PAGES/clinician-edit-form/clinician-edit-form.component';
import {
  CptMasterEditFormComponent,
  CptMasterEditFormModule,
} from '../../POP-UP_PAGES/cpt-master-edit-form/cpt-master-edit-form.component';
import {
  SingleClaimDetailsComponent,
  SingleClaimDetailsModule,
} from '../../REPORT POPUP PAGES/single-claim-details/single-claim-details.component';
import { FormPopupModule } from 'src/app/components';
@Component({
  selector: 'app-department-wise-expense',
  templateUrl: './department-wise-expense.component.html',
  styleUrl: './department-wise-expense.component.scss',
  providers: [ReportService, ReportEngineService],
})
export class DepartmentWiseExpenseComponent {
  @ViewChild(DxDataGridComponent, { static: true })
  dataGrid: DxDataGridComponent;

  @ViewChild(DxTreeViewComponent, { static: false })
  treeView: DxTreeViewComponent;

  @ViewChild('detailsdatagrid', { static: false })
  detailsGrid!: DxDataGridComponent;

  @ViewChild(CptMasterEditFormComponent, { static: false })
  CptEditFormComponent: CptMasterEditFormComponent;

  @ViewChild(ClinicianEditFormComponent, { static: false })
  clinicianEditComponent: ClinicianEditFormComponent;

  @ViewChild('lookup', { static: false }) lookup: DxLookupComponent;

  //=================DataSource for data Grid Table========
  dataGrid_DataSource: DataSource<any>;
  columnsConfig: any; //==============Column data storing variable
  //================Variables for Storing DataSource========

  monthDataSource: { name: string; value: any }[];
  years: number[] = [];

  //================Variables for Storing selected Parameters========
  Facility_Value: any;
  Facility_DataSource: any;
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

  //============Custom close button for drilldown popup============

  userID: any;
  DetailData: any;
  DetailDataCaption: any;
  FilteredDetailData: any;
  DetailsColumns: any;
  DetailsummaryColumns: any;
  originalData: any[] = [];
  originalDetailData: any[] = [];
  SubDetailData: any;
  SubDetailDataCaption: any;
  FilteredSubDetailData: any;
  SubDetailsColumns: any;
  SubDetailsummaryColumns: any;
  ClaimDetailsummaryColumns: any;
  LedgerDetailsummaryColumns:any;

  // LEVEL 2 (ExpenseType)
  headerDataCache: Map<number, any[]> = new Map();

  // LEVEL 3 (ExpenseType + FinanceDepartment)
  detailDataCache: Map<string, any[]> = new Map();

  // LEVEL 4 (ExpenseType + Dept + EntryDept + CostBucket)
  subDetailDataCache: Map<string, any[]> = new Map();

  // LEVEL 5 (FinanceEntry)
  ledgerDataCache: Map<number, any[]> = new Map();

  DepartmentID: any;

  selectedCptCodeData: any;
  selectedClinicianData: any;
  isCptEditFormPopupOpened: boolean = false;
  isEditClinicianPopupOpened: boolean = false;

  ClaimDetailData: any[] = [];
  ClaimColumns: any;

  LedgerDetailData: any[] = [];
  LedgerColumns: any;

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

  //===== Show and Hide Search parameters  ====
  toggleContent() {
    this.isContentVisible = !this.isContentVisible;
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
      return; // Optional: if no further processing needed
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
  //===== Get search parameters dropdown values =====
  get_searchParameters_Dropdown_Values() {
    // this.loadingVisible = true;
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

  //==== Fetch DataSource For The Datagrid Table ======
  async get_Datagrid_DataSource() {
    const formData = {
      FacilityID: this.Facility_Value.join(','),
      PeriodFrom: this.reportengine.formatDate(this.From_Date_Value),
      PeriodTO: this.reportengine.formatDate(this.To_Date_Value),
      Level :1
    };

    this.isContentVisible = false;
    this.dataGrid.instance.beginCustomLoading('Loading...');
    try {
      const response: any = await this.service
        .fetch_Department_Wise_Expense_Data(formData)
        .toPromise();
      if (response.flag === '1') {
        this.isEmptyDatagrid = false;

        // this.columndata = response.header.ReportColumns;

        // const userLocale = navigator.language || 'en-US';

        // this.summaryColumnsData = this.generateSummaryColumns(
        //   response.header.ReportColumns
        // );

        // this.columnsConfig = this.generateColumnsConfig(
        //   response.header.ReportColumns,
        //   userLocale
        // );

        // this.originalData = response.header.ReportData; // store full data
        // this.dataGrid_DataSource = new DataSource<any>({
        //   load: () => Promise.resolve(this.originalData),
        // });

        this.columndata = response.costcentre.ReportColumns;

        this.summaryColumnsData = this.generateSummaryColumns(
          response.costcentre.ReportColumns
        );

        this.columnsConfig = this.generateColumnsConfig(
          response.costcentre.ReportColumns,
          navigator.language || 'en-US'
        );

        this.originalData = response.costcentre.ReportData;

        this.dataGrid_DataSource = new DataSource<any>({
          load: () => Promise.resolve(this.originalData),
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


  
  // ======= main datagrid row expanding function =======
  // async headerDataRowExpanding(e: any) {
  //   const expandedKey = e.key;
  //   this.DepartmentID = expandedKey.DepartmentID || e.data.DepartmentID;
  //   this.FilteredDetailData = [];

  //   // === Collapse all expanded rows first ===
  //   const visibleRows = this.dataGrid.instance.getVisibleRows();
  //   visibleRows.forEach((row) => {
  //     if (row.rowType === 'data' && row.isExpanded) {
  //       this.dataGrid.instance.collapseRow(row.key);
  //     }
  //   });

  //   // === Show loading indicator ===
  //   this.dataGrid.instance.beginCustomLoading('Loading...');

  //   // === Use cache if available ===
  //   if (this.detailDataCache[this.DepartmentID]) {
  //     this.FilteredDetailData = this.detailDataCache[this.DepartmentID];
  //     this.originalDetailData = [...this.FilteredDetailData];
  //     this.dataGrid.instance.endCustomLoading();
  //   } else {
  //     const formData = {
  //       FacilityID: this.Facility_Value.join(','),
  //       PeriodFrom: this.reportengine.formatDate(this.From_Date_Value),
  //       PeriodTO: this.reportengine.formatDate(this.To_Date_Value),
  //       DepartmentID: this.DepartmentID,
  //       FinanceEntryID: 0,
  //     };

  //     try {
  //       const response: any = await this.service
  //         .fetch_Department_Wise_Expense_Data(formData)
  //         .toPromise();

  //       if (response.flag === '1') {
  //         const userLocale = navigator.language || 'en-US';

  //         this.DetailDataCaption = response.detail.ReportID;
  //         this.DetailsummaryColumns = this.generateSummaryColumns(
  //           response.detail.ReportColumns
  //         );
  //         this.DetailsColumns = this.generateColumnsConfig(
  //           response.detail.ReportColumns,
  //           userLocale
  //         );

  //         this.FilteredDetailData = response.detail.ReportData;
  //         this.originalDetailData = [...this.FilteredDetailData];

  //         // Cache result
  //         this.detailDataCache[this.DepartmentID] = [
  //           ...this.FilteredDetailData,
  //         ];
  //       } else {
  //         notify({ message: response.message }, 'error');
  //         this.dataGrid.instance.endCustomLoading();
  //         return;
  //       }
  //     } catch (error) {
  //       notify({ message: 'Error loading data' }, 'error');
  //       this.dataGrid.instance.endCustomLoading();
  //       return;
  //     } finally {
  //       this.dataGrid.instance.endCustomLoading();
  //     }
  //   }

  //   // // === Set only the expanded row in grid ===
  //   // const expandedRow = this.originalData.find(
  //   //   (row) => row.DepartmentID === this.DepartmentID
  //   // );
  //   // this.dataGrid_DataSource = new DataSource<any>({
  //   //   load: () => Promise.resolve([expandedRow]),
  //   // });

  //   // // === Scroll to the expanded row ===
  //   // setTimeout(() => {
  //   //   const rowIndex = this.dataGrid.instance.getRowIndexByKey(e.key);
  //   //   if (rowIndex >= 0) {
  //   //     const scrollable = (this.dataGrid.instance as any).getScrollable();
  //   //     const rowElements: any = this.dataGrid.instance.getRowElement(rowIndex);
  //   //     if (scrollable && rowElements?.length > 0) {
  //   //       const topOffset = rowElements[0].offsetTop;
  //   //       scrollable.scrollTo({ top: topOffset });
  //   //     }
  //   //   }
  //   // }, 0);
  // }

  async headerDataRowExpanding(e: any) {
  const rowData = e.data || e.key;

  if (!rowData || !rowData.ExpenseTypeID) {
    console.warn('ExpenseTypeID not found', e);
    return;
  }

  const ExpenseTypeID = rowData.ExpenseTypeID;

  // 🔹 USE CACHE
  // if (this.headerDataCache[ExpenseTypeID]) {
  //   this.FilteredDetailData = this.headerDataCache[ExpenseTypeID];
  //   return;
  // }

  this.FilteredDetailData = [];
  this.dataGrid.instance.beginCustomLoading('Loading...');

  const headerGridInstance = e.component; // ✅ CURRENT DETAIL GRID

  // 🔴 CLOSE all other expanded rows
  const visibleRows = headerGridInstance.getVisibleRows();
  visibleRows.forEach((row) => {
    if (
      row.rowType === 'data' &&
      row.isExpanded &&
      row.key !== e.key
    ) {
      headerGridInstance.collapseRow(row.key);
    }
  });

  const formData = {
    FacilityID: this.Facility_Value.join(','),
    PeriodFrom: this.reportengine.formatDate(this.From_Date_Value),
    PeriodTO: this.reportengine.formatDate(this.To_Date_Value),
    ExpenseTypeID: ExpenseTypeID,
    Level :2
  };

  try {
    const response: any = await this.service
      .fetch_Department_Wise_Expense_Data(formData)
      .toPromise();

    if (response.flag === '1' && response.header) {

      this.DetailsColumns = this.generateColumnsConfig(
        response.header.ReportColumns,
        navigator.language || 'en-US'
      );

      this.DetailsummaryColumns = this.generateSummaryColumns(
        response.header.ReportColumns
      );

      this.FilteredDetailData = response.header.ReportData;
      this.originalDetailData = [...this.FilteredDetailData];

      // ✅ CACHE IT
      // this.headerDataCache[ExpenseTypeID] = [
      //   ...this.FilteredDetailData
      // ];

    } else {
      notify(response.message || 'No data found', 'error');
    }
  } catch {
    notify('Error loading header data', 'error');
  } finally {
    this.dataGrid.instance.endCustomLoading();
  }
}


  // ======= main datagrid row collapsing function =======
  onRowCollapsing(e: any) {
    // Reset full data source
    this.dataGrid_DataSource = new DataSource<any>({
      load: () => Promise.resolve(this.originalData),
    });
  }

  // ======= details datagrid row expanding function =======
  // async detailDataRowExpanding(e: any) {
  //   const FinanceEntryID = e.key.FinanceEntryID || e.data.FinanceEntryID;

  //   // Show only clicked detail row
  //   const expandedRow = this.originalDetailData.find(
  //     (row) => row.FinanceEntryID === FinanceEntryID
  //   );
  //   this.FilteredDetailData = [expandedRow];
  //   this.FilteredSubDetailData = [];

  //   // Show loading
  //   this.dataGrid.instance?.beginCustomLoading?.('Loading...');

  //   // Use cached sub-detail
  //   if (this.subDetailDataCache[FinanceEntryID]) {
  //     this.FilteredSubDetailData = this.subDetailDataCache[FinanceEntryID];
  //     this.dataGrid.instance?.endCustomLoading?.();
  //     return;
  //   }

  //   const formData = {
  //     FacilityID: this.Facility_Value.join(','),
  //     PeriodFrom: this.reportengine.formatDate(this.From_Date_Value),
  //     PeriodTO: this.reportengine.formatDate(this.To_Date_Value),
  //     DepartmentID: this.DepartmentID,
  //     FinanceEntryID: FinanceEntryID,
  //     Level :3
  //   };

  //   try {
  //     const response: any = await this.service
  //       .fetch_Department_Wise_Expense_Data(formData)
  //       .toPromise();

  //     if (response.flag === '1') {
  //       const userLocale = navigator.language || 'en-US';

  //       this.SubDetailDataCaption = response.subdetail.ReportID;
  //       this.SubDetailsummaryColumns = this.generateSummaryColumns(
  //         response.subdetail.ReportColumns
  //       );
  //       this.SubDetailsColumns = this.generateColumnsConfig(
  //         response.subdetail.ReportColumns,
  //         userLocale
  //       );

  //       this.FilteredSubDetailData = response.subdetail.ReportData;

  //       // Cache it
  //       this.subDetailDataCache[FinanceEntryID] = [
  //         ...this.FilteredSubDetailData,
  //       ];
  //     } else {
  //       notify({ message: response.message }, 'error');
  //     }
  //   } catch (error) {
  //     notify({ message: 'Error loading sub-detail data' }, 'error');
  //   } finally {
  //     this.dataGrid.instance?.endCustomLoading?.();
  //   }
  // }

  async detailDataRowExpanding(e: any) {
  const rowData = e.data || e.key;

  const ExpenseTypeID = rowData?.ExpenseTypeID 
  const FinanceDepartmentID = rowData?.FinanceDepartmentID 

  const cacheKey = `${ExpenseTypeID}_${FinanceDepartmentID}`;

  // ✅ CACHE HIT
  if (this.detailDataCache.has(cacheKey)) {
    this.FilteredSubDetailData = this.detailDataCache.get(cacheKey)!;
    return;
  }



  const detailGridInstance = e.component; // ✅ CURRENT DETAIL GRID

  // 🔴 CLOSE all other expanded rows
  const visibleRows = detailGridInstance.getVisibleRows();
  visibleRows.forEach((row) => {
    if (
      row.rowType === 'data' &&
      row.isExpanded &&
      row.key !== e.key
    ) {
      detailGridInstance.collapseRow(row.key);
    }
  });

  const formData = {
    FacilityID: this.Facility_Value.join(','),
    PeriodFrom: this.reportengine.formatDate(this.From_Date_Value),
    PeriodTO: this.reportengine.formatDate(this.To_Date_Value),
    ExpenseTypeID: ExpenseTypeID,
    FinanceDepartmentID: FinanceDepartmentID,
    Level:3
  };

  this.FilteredSubDetailData = [];
  this.dataGrid.instance.beginCustomLoading('Loading...');

  try {
    const response: any = await this.service
      .fetch_Department_Wise_Expense_Data(formData)
      .toPromise();
    
    if (response.flag === '1' && response.detail) {
      this.SubDetailsColumns = this.generateColumnsConfig(
        response.detail.ReportColumns,
        navigator.language || 'en-US'
      );

      this.SubDetailsummaryColumns = this.generateSummaryColumns(
        response.detail.ReportColumns
      );

      this.FilteredSubDetailData = response.detail.ReportData;

      // ✅ CACHE IT
      this.detailDataCache.set(cacheKey, [...this.FilteredSubDetailData]);
    }
  } catch {
    notify('Error loading detail data', 'error');
  } finally {
    this.dataGrid.instance.endCustomLoading();
  }
}



async subDetailRowExpanding(e: any) {
  const rowData = e.data || e.key;


  const ExpenseTypeID = rowData?.ExpenseTypeID
  const FinanceDepartmentID = rowData?.FinanceDepartmentID
  const FinanceEntryDepartmentID = rowData?.FinanceEntryDepartmentID
  const CostBucketID = rowData?.CostBucketID

  // // 🔴 CLOSE other expanded claims
  // this.collapseOtherRows(grid, e.key);

  const cacheKey = `${ExpenseTypeID}_${FinanceDepartmentID}_${FinanceEntryDepartmentID}_${CostBucketID}`;

  // ✅ CACHE HIT
  if (this.subDetailDataCache.has(cacheKey)) {
    this.LedgerDetailData = this.subDetailDataCache.get(cacheKey)!;
    return;
  }

  const subDetailGrid = e.component;

  // 🔴 Close other expanded Level-4 rows (important)
  const visibleRows = subDetailGrid.getVisibleRows();
  visibleRows.forEach((row) => {
    if (
      row.rowType === 'data' &&
      row.isExpanded &&
      row.key !== e.key
    ) {
      subDetailGrid.collapseRow(row.key);
    }
  });

  // 🔹 Clear previous claim data
  this.ClaimDetailData = [];
  this.ClaimColumns = null;

  this.LedgerDetailData = [];
  this.LedgerColumns = null;

  const formData = {
    FacilityID: this.Facility_Value.join(','),
    PeriodFrom: this.reportengine.formatDate(this.From_Date_Value),
    PeriodTO: this.reportengine.formatDate(this.To_Date_Value),
    ExpenseTypeID: ExpenseTypeID,
    FinanceDepartmentID: FinanceDepartmentID,
    FinanceEntryDepartmentID : FinanceEntryDepartmentID,
    CostBucketID : CostBucketID,
    Level: 4
  };

  this.dataGrid.instance.beginCustomLoading('Loading...');

  try {
    const response: any = await this.service
      .fetch_Department_Wise_Expense_Data(formData)
      .toPromise();

    if (response.flag === '1' && response.ledger) {
       // 🔹 Bind Level-4 columns
      this.LedgerColumns = this.generateColumnsConfig(
        response.ledger.ReportColumns,
        navigator.language || 'en-US'
      );

      this.LedgerDetailsummaryColumns = this.generateSummaryColumns(
        response.ledger.ReportColumns
      );


      // 🔹 Bind Level-4 data
      this.LedgerDetailData = response.ledger.ReportData;

      // ✅ CACHE
      this.subDetailDataCache.set(cacheKey, [...this.LedgerDetailData]);
    }
  } finally {
    this.dataGrid.instance.endCustomLoading();
  }
}


async ledgerRowExpanding(e: any) {
  const rowData = e.data || e.key;


  const FinanceEntryID = rowData?.FinanceEntryID; // 👈 usually needed

  // ✅ CACHE HIT
  if (this.ledgerDataCache.has(FinanceEntryID)) {
    this.ClaimDetailData = this.ledgerDataCache.get(FinanceEntryID)!;
    return;
  }


  // 🔴 Close other expanded Level-5 rows
  const grid = e.component;
  grid.getVisibleRows().forEach(r => {
    if (r.rowType === 'data' && r.isExpanded && r.key !== e.key) {
      grid.collapseRow(r.key);
    }
  });

  this.ClaimDetailData = [];
  this.dataGrid.instance.beginCustomLoading('Loading...');

  const formData = {
    FacilityID: this.Facility_Value.join(','),
    PeriodFrom: this.reportengine.formatDate(this.From_Date_Value),
    PeriodTO: this.reportengine.formatDate(this.To_Date_Value),
    FinanceEntryID : FinanceEntryID,
    Level: 5   // ✅ NEW LEVEL
  };

  try {
    const response: any = await this.service
      .fetch_Department_Wise_Expense_Data(formData)
      .toPromise();

    if (response.flag === '1' && response.subdetail) {
      this.ClaimColumns = this.generateColumnsConfig(
        response.subdetail.ReportColumns,
        navigator.language || 'en-US'
      );

      this.ClaimDetailsummaryColumns = this.generateSummaryColumns(
        response.subdetail.ReportColumns
      );

      this.ClaimDetailData = response.subdetail.ReportData;
      // ✅ CACHE
      this.ledgerDataCache.set(FinanceEntryID, [...this.ClaimDetailData]);
    }
  } catch {
    notify('Error loading Level 5 data', 'error');
  } finally {
    this.dataGrid.instance.endCustomLoading();
  }
}




  // ======= details datagrid row collapsing function =======
  onDetailRowCollapsing(e: any) {
    this.FilteredDetailData = [...this.originalDetailData];
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

      if (column.Name === 'Qty_Weight') {
        columnFormat = {
          formatter: this.qtyWeightFormatter,
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

  qtyWeightFormatter = (value: any) => {
    if (value === 0 || value === '0' || value == null) return '';
    return Number(value).toFixed(2);
  };

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
  //============Show Filter Row==========================
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

  //======= Exporting Function =======
  onExporting(event: any) {
    const fileName = 'department wise expense';
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
  declarations: [DepartmentWiseExpenseComponent],
})
export class DepartmentWiseExpenseModule {}
