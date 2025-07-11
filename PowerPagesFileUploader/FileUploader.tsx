import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { PrimaryButton, DefaultButton, Modal, IconButton, IIconProps, DetailsList, IColumn, DetailsListLayoutMode, Spinner, SpinnerSize, Dialog, DialogType, DialogFooter, IDialogContentProps, IDialogStyles } from "@fluentui/react";
import { BlobServiceClient } from "@azure/storage-blob";
import { renderAsync } from "docx-preview";
import "./FileUploader.css";
import { getSasToken } from "./SasTokenRetriever";

/* eslint-disable @typescript-eslint/no-explicit-any */
interface IFileUploaderProps {
  context: ComponentFramework.Context<any>;
  notifyOutputChanged: () => void;
  height: string;
  width: string;
  blobAccountName: string;
  blobContainerName: string;
  sasTokenEnvironmentVariable: string;
  powerPagesUrl: string;
}
/* eslint-enable @typescript-eslint/no-explicit-any */

const cancelIcon: IIconProps = { iconName: 'Cancel' };
const deleteIcon: IIconProps = { iconName: 'Delete' };
const downloadIcon: IIconProps = { iconName: 'Download' };

export const FileUploader: React.FC<IFileUploaderProps> = ({
  context,
  notifyOutputChanged,
  height,
  width,
  blobAccountName,
  blobContainerName,
  sasTokenEnvironmentVariable,
  powerPagesUrl
}) => {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [dragging, setDragging] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isPdf, setIsPdf] = useState(false);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [fileDetails, setFileDetails] = useState<any[]>([]);
// eslint-enable-next-line @typescript-eslint/no-explicit-any
  const [currentUserId, setCurrentUserId] = useState<string | null>(null);
  const [recordId, setRecordId] = useState<string | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [sasToken, setSasToken] = useState<string | null>(null);
  const [deleteDialogVisible, setDeleteDialogVisible] = useState<boolean>(false);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [fileToDelete, setFileToDelete] = useState<any | null>(null);
  // eslint-enable-next-line @typescript-eslint/no-explicit-any
  const previewRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);  

  useEffect(() => {
    const fetchSasToken = async () => {
      try {
        setLoading(true);
        const token = await getSasToken(context, sasTokenEnvironmentVariable, powerPagesUrl);
        setSasToken(token);
        setLoading(false);
      } catch (error) {
        console.error("Error retrieving SAS token:", error);
        setLoading(false);
      }
    };

    fetchSasToken();
  }, [sasTokenEnvironmentVariable, powerPagesUrl]);

  useEffect(() => {
    if (sasToken) {
        getCurrentUserId().then((userId) => {
            setCurrentUserId(userId);
            if (userId) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const mdarecordId = (context.mode as any).contextInfo.entityId;
                console.log(`mdarecordId v4: ${mdarecordId}`);
                // eslint-enable-next-line @typescript-eslint/no-explicit-any
                setRecordId(mdarecordId);
                fetchBlobFiles(userId, mdarecordId);
              }            
          });
    }
  }, [sasToken]);

  const getCurrentUserId = async () => {
    try {
        let userId = context.userSettings.userId;
        console.log(`Dataverse userId: ${userId}`);
        if(userId == null)
        {
            console.log(`powerPagesUrl: ${powerPagesUrl}`);
            const response = await fetch(`${powerPagesUrl}/_api/contacts?$select=contactid`, {
                method: 'GET',
                headers: {
                  'Content-Type': 'application/json'
                }
              });
        
              const data = await response.json();
              
              if (data && data.value && data.value.length > 0) {
                userId = data.value[0].contactid;
                console.log(`Power Pages userId: ${userId}`);
              }
        }
        return userId;
      } catch (error) {
        console.error("Error fetching user ID:", error);
        return null;
      }
    };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      setSelectedFile(event.target.files[0]);
      setIsPdf(event.target.files[0].type === "application/pdf");
    }
  };

  const handleDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(true);
  };

  const handleDragLeave = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(false);
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(false);
    if (event.dataTransfer.files && event.dataTransfer.files.length > 0) {
      setSelectedFile(event.dataTransfer.files[0]);
      setIsPdf(event.dataTransfer.files[0].type === "application/pdf");
      event.dataTransfer.clearData();
    }
  };

  const handleUpload = async () => {
    if (selectedFile && sasToken) {
      const accountName = blobAccountName;
      const containerName = blobContainerName;

      const blobServiceClient = new BlobServiceClient(
        `https://${accountName}.blob.core.windows.net?${sasToken}`
      );
      const containerClient = blobServiceClient.getContainerClient(containerName);
      const blockBlobClient = containerClient.getBlockBlobClient(selectedFile.name);    

      const metadata: { [key: string]: string } = {
        uploadedby: currentUserId || "",
        status: "Uploaded" // Default status
      };
      if (recordId) {
        metadata.recordid = recordId; // Add recordid only if it's available
      }

      try {
        await blockBlobClient.uploadBrowserData(selectedFile, {
          blobHTTPHeaders: {
            blobContentType: selectedFile.type
          },
          metadata: metadata
        });
        alert("File uploaded successfully!");
        fetchBlobFiles(currentUserId, recordId);
        notifyOutputChanged();
        setSelectedFile(null);
        if (fileInputRef.current) {
            fileInputRef.current.value = "";
        }        
      } catch (error) {
        console.error("Error uploading file:", error);
        alert("File upload failed!");
      }
    }
  };

  const fetchBlobFiles = async (userId: string | null, recordId: string | null) => {
    if (!sasToken && !userId) return;

    setLoading(true);

    const accountName = blobAccountName;
    const containerName = blobContainerName;

    const blobServiceClient = new BlobServiceClient(
      `https://${accountName}.blob.core.windows.net?${sasToken}`
    );
    const containerClient = blobServiceClient.getContainerClient(containerName);
    console.log(`fetchBlobFiles currentUserId: ${userId}`);
    console.log(`recordId: ${recordId}`);
    try {
      const listBlobsResponse = containerClient.listBlobsFlat();
      const files = [];
      for await (const blob of listBlobsResponse) {
        const blockBlobClient = containerClient.getBlockBlobClient(blob.name);
        const blobProperties = await blockBlobClient.getProperties();

        const metadata = blobProperties.metadata || {};

        console.log(`Blob metadata for ${blob.name}: ${JSON.stringify(metadata)}`);
        console.log(`Filtering by userId: ${userId}, recordId: ${recordId}`);

        if(recordId)
        {
            console.log(`MDA`);
            //if (metadata.uploadedby === userId && metadata.recordid === recordId) {
            if (metadata.recordid === recordId) {
                files.push({
                  key: blob.name,
                  name: blob.name,
                  size: ((blobProperties.contentLength ?? 0) / 1024).toFixed(2), // Divide by 1024 to get size in KB and format to 2 decimals
                  modified: blobProperties.lastModified?.toLocaleString(),
                  status: metadata.status || "N/A",
                  recordid: metadata.recordid || "N/A"
                });
            }
        }
        else {
            console.log(`Power Pages`);
            //if (metadata.uploadedby === userId) {
            if (metadata.recordid === recordId) {
                files.push({
                    key: blob.name,
                    name: blob.name,
                    size: ((blobProperties.contentLength ?? 0) / 1024).toFixed(2), // Divide by 1024 to get size in KB and format to 2 decimals
                    modified: blobProperties.lastModified?.toLocaleString(),
                    status: metadata.status || "N/A",
                    recordid: metadata.recordid || "N/A"
                });
            }
        }
      }
      setFileDetails(files);
    } catch (error) {
      console.error("Error fetching blob files:", error);
    }
    finally {
        setLoading(false);
    }
  };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const confirmDelete = (item: any) => {
      // eslint-enable-next-line @typescript-eslint/no-explicit-any
    setFileToDelete(item);
    setDeleteDialogVisible(true);
  };

  const handleDelete = async () => {
    if (!sasToken || !fileToDelete) return;

    const accountName = blobAccountName;
    const containerName = blobContainerName;

    const blobServiceClient = new BlobServiceClient(
      `https://${accountName}.blob.core.windows.net?${sasToken}`
    );
    const containerClient = blobServiceClient.getContainerClient(containerName);
    const blockBlobClient = containerClient.getBlockBlobClient(fileToDelete.name);

    try {
      await blockBlobClient.delete();
      alert("File deleted successfully!");
      fetchBlobFiles(currentUserId, recordId);
      notifyOutputChanged();
      setFileToDelete(null);
      setDeleteDialogVisible(false);
    } catch (error) {
      console.error("Error deleting file:", error);
      alert("File deletion failed!");
    }
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const handleDownload = async (item: any) => {
  // eslint-enable-next-line @typescript-eslint/no-explicit-any
    if (!sasToken) return;

    const accountName = blobAccountName;
    const containerName = blobContainerName;

    const blobServiceClient = new BlobServiceClient(
      `https://${accountName}.blob.core.windows.net?${sasToken}`
    );
    const containerClient = blobServiceClient.getContainerClient(containerName);
    const blockBlobClient = containerClient.getBlockBlobClient(item.name);

    try {
      const downloadResponse = await blockBlobClient.download();
      const blob = await downloadResponse.blobBody;
      if (blob) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = item.name;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }
    } catch (error) {
      console.error("Error downloading file:", error);
      alert("File download failed!");
    }
  };

  const columns: IColumn[] = [
    { key: 'column1', name: 'Name', fieldName: 'name', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column2', name: 'Size (KB)', fieldName: 'size', minWidth: 50, maxWidth: 100, isResizable: true },
    { key: 'column3', name: 'Modified', fieldName: 'modified', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column4', name: 'Status', fieldName: 'status', minWidth: 100, maxWidth: 100, isResizable: true },
    {
      key: 'column5',
      name: 'Actions',
      fieldName: 'actions',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
// eslint-disable-next-line @typescript-eslint/no-explicit-any
      onRender: (item: any) => (
// eslint-enable-next-line @typescript-eslint/no-explicit-any
        <div>
          <IconButton
            iconProps={downloadIcon}
            title="Download"
            ariaLabel="Download"
            onClick={() => handleDownload(item)}
          />
          <IconButton
            iconProps={deleteIcon}
            title="Delete"
            ariaLabel="Delete"
            onClick={() => confirmDelete(item)}
          />
        </div>
      )
    }
  ];

  const toggleModal = () => {
    setIsModalOpen(!isModalOpen);
    if (!isModalOpen && selectedFile && (selectedFile.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || selectedFile.type === "application/msword")) {
      const reader = new FileReader();
      reader.onload = async (event) => {
        const arrayBuffer = event.target?.result as ArrayBuffer;
        if (arrayBuffer && previewRef.current) {
          try {
            await renderAsync(arrayBuffer, previewRef.current);
          } catch (error) {
            console.error("Error rendering document:", error);
            previewRef.current.innerHTML = "Failed to load preview.";
          }
        }
      };
      reader.readAsArrayBuffer(selectedFile);
    }
  };

  const renderFilePreview = () => {
    if (selectedFile) {
      const fileURL = URL.createObjectURL(selectedFile);
      if (selectedFile.type.startsWith("image/")) {
        return <img src={fileURL} alt="File preview" style={{ maxWidth: "100%", maxHeight: "400px" }} />;
      } else if (selectedFile.type.startsWith("video/")) {
        return <video src={fileURL} controls style={{ maxWidth: "100%", maxHeight: "400px" }} />;
      } else if (selectedFile.type === "application/pdf") {
        return <iframe src={fileURL} style={{ width: "1400px", height: "650px" }} />;
      } else if (
        selectedFile.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
        selectedFile.type === "application/msword"
      ) {
        return <div ref={previewRef} />;
      } else {
        return <div>No preview available for this file type.</div>;
      }
    }
    return null;
  };

  const dialogStyles: Partial<IDialogStyles> = {
    main: {
      maxWidth: 450,
      width: '90%',
      textAlign: 'center'
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'row', height, width }}>
      <div
        className={`file-uploader ${dragging ? "dragging" : ""}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        style={{ height, width: '50%', border: '2px dashed #ccc', borderRadius: '4px', textAlign: 'center', padding: '20px' }}
      >
        <input type="file" onChange={handleFileChange} ref={fileInputRef} style={{ display: "none" }} id="fileInput" />
        <label htmlFor="fileInput" className="file-input-label">
          Click to select a file or drag it here
        </label>
        {selectedFile && <div>Selected file: {selectedFile.name}</div>}
        <div style={{ marginTop: "10px" }}>
          <PrimaryButton onClick={handleUpload} disabled={!selectedFile}>
            Upload
          </PrimaryButton>
          {selectedFile && (
            <DefaultButton onClick={toggleModal} style={{ marginLeft: "10px" }}>
              Preview
            </DefaultButton>
          )}
        </div>
        <Modal
          isOpen={isModalOpen}
          onDismiss={toggleModal}
          isBlocking={false}
          containerClassName={isPdf ? "modal-container-large" : ""}
        >
          <div className="modal-header">
            <span>File Preview</span>
            <IconButton
              iconProps={cancelIcon}
              ariaLabel="Close popup modal"
              onClick={toggleModal}
            />
          </div>
          <div className={`modal-content ${isPdf ? "pdf-preview" : ""}`}>
            {renderFilePreview()}
          </div>
        </Modal>
      </div>
      <div style={{ height, width: '50%', paddingLeft: '10px' }}>
        {loading ? (
          <Spinner size={SpinnerSize.large} label="Loading files..." />
        ) : (
          <DetailsList
            items={fileDetails}
            columns={columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
          />
        )}
      </div>

      <Dialog
        hidden={!deleteDialogVisible}
        onDismiss={() => setDeleteDialogVisible(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Confirm Deletion',
          closeButtonAriaLabel: 'Close',
          subText: 'Are you sure you want to delete this file?'
        }}
        styles={dialogStyles}
      >
        <DialogFooter>
          <PrimaryButton onClick={handleDelete} text="Yes" />
          <DefaultButton onClick={() => setDeleteDialogVisible(false)} text="No" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
