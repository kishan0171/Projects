
var mainApp = angular.module('reportApp', ['ngTouch', 'ngAnimate', 'ui.grid', 'ui.grid.pagination', 'ui.grid.exporter']);


mainApp.controller("reportController", ['$scope', '$http', function ($scope, $http) {
    var globMarketType = "GM";

    $http({
        url: "../_WebService/WebService1.asmx/GetDataDup",
        dataType: 'json',
        method: 'POST',
        data: "{}",
        headers: {
            "Content-Type": "application/json"
        }
    }).success(function (response) {
        OrderList = response.d;
        //alert(response.d);
        $scope.gridOptions1.data = jQuery.parseJSON(OrderList);
        //$scope.gridOptions1.data = OrderList;
    })
    .error(function (response) {
        alert(response.data);
    });

    //var inputTxtManualAllocQty = '<div><form name="inputForm"><input type="text" class="k-input k-textbox" ' +
    //    'ng-model="row.entity.ManualOrdAllocatedQty" value="{{row.entity.ManualOrdAllocatedQty}}" ' +
    //    'style="height: 25px; width: 120px;" /> </form></div>';

    //var buttonLink = '<div> <button type="button" name="{{row.entity.F1}}" onclick="EditData(row.entity)" /> </div>'

    var buttonLink = "<div class='ui-grid-cell-contents'>" +
   "<button type='button' class='btn btn-xs btn-primary' ng-click='grid.appScope.editRow(row)'>" +
   " Edit " +
   "</button>" +
   "</div>";

    //var inputTxtManualAllocQty = '<div><a name="inputForm"><input type="text" class="k-input k-textbox" ' +
    //    'ng-model="row.entity.ManualOrdAllocatedQty" value="{{row.entity.ManualOrdAllocatedQty}}" ' +
    //    'style="height: 25px; width: 120px;" /> </a></div>';

    $scope.gridOptions1 = {
        paginationPageSizes: [5, 10],
        paginationPageSize: 10,
        enableColumnResizing: true,
        enableGridMenu: true,
        enableCellEditOnFocus: false,
        enableSelectAll: true,
        enableFiltering: true,
        exporterCsvFilename: 'orderBook.csv',
        exporterPdfDefaultStyle: { fontSize: 9 },
        exporterPdfTableStyle: { margin: [10, 10, 5, 10] },
        exporterPdfTableHeaderStyle: { fontSize: 10, bold: true, italics: true, color: 'white', background: 'black' },
        exporterPdfHeader: { text: "Order Book", style: 'headerStyle', alignment: 'center' },
        exporterPdfFooter: function (currentPage, pageCount) {
            return { text: currentPage.toString() + ' of ' + pageCount.toString(), style: 'footerStyle' };
        },
        exporterPdfCustomFormatter: function (docDefinition) {
            docDefinition.styles.headerStyle = { fontSize: 22, bold: true };
            docDefinition.styles.footerStyle = { fontSize: 10, bold: true };
            return docDefinition;
        },
        exporterPdfOrientation: 'landscape',
        exporterPdfPageSize: 'A3',
        exporterPdfMaxGridWidth: 950,
        exporterCsvLinkElement: angular.element(document.querySelectorAll(".custom-csv-link-location")),
        onRegisterApi: function (gridApi) {
            $scope.gridApi = gridApi;
        },

        columnDefs: [
        	             { name: 'F1', displayName: 'F1', width: 100 },
        		         { name: 'F2', displayName: 'Address', width: 150 },
                         { name: 'F3', displayName: 'Street', width: 150 },
        		         { name: 'F4', displayName: 'Side', width: 100 },
        		         { name: 'F5', displayName: 'Site', width: 100 },
        		         { name: 'F6', displayName: 'Species', width: 100 },
        		         { name: 'F7', displayName: 'DBH', width: 100 },
                         { name: 'F8', displayName: 'Condition', width: 100 },
                         {
                             name: 'Edit', displayName: 'Edit (editable)', width: 150,
                             cellTemplate: buttonLink
                         }
        ]
    };

    $scope.editRow = function (row) {
        $scope.EditDataRecord = row.entity;
    }

    $scope.Submit = function (row) {
        if (row.F7 < 0) {
            alert("Enter more then Zero !!!")
        }
    }


}]);