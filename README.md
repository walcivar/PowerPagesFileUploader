The goal of this PCF is to allow an external user through Power Pages to upload documents by drag and drop with the option to preview the document before uploading it so that the user can be sure of the document being uploaded.

This PCF also works for Model Driven Apps so internal users can do the same.

The document repository for this PCF is Azure blob storage so first you will have to configure this service if you want to use this PCF.


1)	Configure Azure blob storage:

Create a storage account:
Create a Container
Configure CORS Settings under the Storage account

![image](https://github.com/walcivar/PowerPagesFileUploader/assets/5630463/cc47742e-494a-41cb-883f-7b51f2ebca96)

Create a Shared access signature

![image](https://github.com/walcivar/PowerPagesFileUploader/assets/5630463/a537ef8d-c1cd-470d-851e-7123e3ef0457)

Click in Generate SAS and connection string and copy the SAS token

![image](https://github.com/walcivar/PowerPagesFileUploader/assets/5630463/2208cecb-0918-4bd5-9d77-61ac030a5b22)

2)	Create a text environment variable in your solution, then copy and paste the created SAS token
3)	Import the managed or unmanaged solution in your environment
4)	Create a text column 

For Model Driven App

5)	Add the PCF to the new column and configure it as the following image, for Height set it as 400 for width set it as 100%, the rest parameters you will have to set your own values:

![image](https://github.com/walcivar/PowerPagesFileUploader/assets/5630463/1807b432-9803-49b3-8b71-bbb1db28ed1b)

That’s it now you can use the PCF in a Model Driven App.

If you want to use it also in Power Pages, do these additional steps:
6)	Beside configuring the PCF in the form you will have to add a code component metadata to the form:

![image](https://github.com/walcivar/PowerPagesFileUploader/assets/5630463/1e9bd7bb-f23c-442a-92fa-96d5d95c6a4c)

7)	Don’t forget to turn on the Webapi for the following tables:
   
![image](https://github.com/walcivar/PowerPagesFileUploader/assets/5630463/bc51cb35-0a0f-43a5-9f7b-29a324b771d3)

8) Create the table permission records for Contact, Environment Variable Definition and Environment Variable Value tables:

![image](https://github.com/walcivar/PowerPagesFileUploader/assets/5630463/53268ffd-fa46-47a7-b1db-99deb8dffa70)
  
9)	Clear the site cache and you are free to go

![image](https://github.com/walcivar/PowerPagesFileUploader/assets/5630463/23fa5ac1-7aa4-403a-89f9-031b2dc0f4f3)

![FileUploaderPCF](https://github.com/walcivar/PowerPagesFileUploader/assets/5630463/f38576f7-3408-4359-914a-6621fe80825d)
