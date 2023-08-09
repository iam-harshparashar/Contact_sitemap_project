import { Component } from '@angular/core';
import { ExcelService } from './excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  public file: File | null = null;
  public records: any[] = [];
  public tableRecords: any[] = [];
  message: string | undefined;
  isSubmitButtonDisabled: boolean = true;

  constructor(private excelService: ExcelService) {}

  // Event handler for when a file is dropped or selected

  onFileDropped(event: any) {

    this.file = event.target.files[0];
    this.message = 'Invalid';

    if (this.file) {

      this.isSubmitButtonDisabled=false;

      this.excelService.parseExcel(this.file).then(async (data) => {
        this.tableRecords = data;
        console.log(data);
        // this.records = data;
        console.log(this.records);
        console.log(this.tableRecords);

        // Loop through each record from the Excel data

        for (var i = 0; i < data.length; i++) {
          var gender;
          if (String(data[i].gendercode).toLowerCase() == 'male') gender = 1;
          else if (String(data[i].gendercode).toLowerCase() == 'female')
            gender = 2;

          // Convert the birthdate to a suitable format for the API or set it to null if not provided

          const birthdate = data[i].birthdate
            ? new Date(data[i].birthdate).toISOString().split('T')[0]
            : null;

          // Retrieve the Account ID associated with the current record's name

          const accountId = await this.retrieveAccountId(String(data[i].name));

          // Prepare the data to be used for creating a new contact record using Web API

          var convertedData = {
            firstname: String(data[i].firstname),
            lastname: String(data[i].lastname),
            emailaddress1: String(data[i].emailaddress1),
            mobilephone: String(data[i].mobilephone),
            birthdate: birthdate,
            gendercode: gender,
            'parentcustomerid_account@odata.bind': `/accounts(${accountId})`,
          };

          // Push the transformed record into the records array

          this.records.push(convertedData);
        }
      });
      this.message = 'Data Imported Successfully!';
    }
  }

  // Asynchronously retrieve the Account ID based on the given accountName

  async retrieveAccountId(accountName: any) {
    try {

      // Use Xrm.WebApi to retrieve the Account record based on the name

      const response = await Xrm.WebApi.retrieveMultipleRecords(
        'account',
        `?$select=accountid&$filter=name eq '${accountName}'`
      );

      if (response.entities.length > 0) {
        // Return the Account ID
        return response.entities[0].accountid;
      } else {
        // Return null if the accountName is not found
        return null;
      }
    } catch (error) {
      // Handle any errors that occur during the retrieval
      console.log(error);
      return null;
    }
  }

  // Process the records array and create new contact records using Web API

  processFile() {
    this.isSubmitButtonDisabled=true;
    for (const entity of this.records) {
      Xrm.WebApi.createRecord('contact', entity).then(
        function (result) {
          // Log the ID of the created contact record
          console.log('Created contact with ID: ' + result.id);
        },
        function (error) {
          // Handle the error
          console.log(error.message);
        }
      );
    }
    this.message = 'Records Created Successfully';
  }
}
