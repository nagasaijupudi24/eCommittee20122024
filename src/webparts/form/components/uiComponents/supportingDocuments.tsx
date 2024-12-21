/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @rushstack/no-new-null */
import * as React from "react";
import { IconButton } from "@fluentui/react";


import styles from "../Form.module.scss";

export interface IUploadFileProps {
  cummulativeError: any;
  typeOfDoc: string;
  onChange: (files: File[] | null, typeOfDoc: string) => void;
  accept?: string;
  maxFileSizeMB: number;
  multiple: boolean;
  
  data: File[];
  errorData: any;
  addtionalData: any[];
}

interface IFileWithError {
  id: string;
  file: File;
  error: string | null;
  cumulativeError: any;
}

interface IUploadFileState {
  selectedFiles: IFileWithError[];
  cummError: string | null;
  errorOfFile: any;
}

const _randomFileIcon = (docType: string): any => {
 
  const docExtension = docType.split(".");
  const fileExtession = docExtension[docExtension.length - 1];

  let doctype = "txt";

  switch (fileExtession.toLocaleLowerCase()) {
    case "accdb":
      doctype = "accdb";
      break;
    case "audio":
      doctype = "audio";
      break;
    case "code":
      doctype = "code";
      break;
    case "csv":
      doctype = "csv";
      break;
    case "docx":
      doctype = "docx";
      break;
    case "doc":
      doctype = "docx";
      break;
    case "dotx":
      doctype = "dotx";
      break;
    case "mpp":
      doctype = "mpp";
      break;
    case "mpt":
      doctype = "mpt";
      break;
    case "model":
      doctype = "model";
      break;
    case "one":
      doctype = "one";
      break;
    case "onetoc":
      doctype = "onetoc";
      break;
    case "potx":
      doctype = "potx";
      break;
    case "ppsx":
      doctype = "ppsx";
      break;
    case "pdf":
      doctype = "pdf";
      break;
    case "photo":
      doctype = "photo";
      break;
    case "pptx":
      doctype = "pptx";
      break;
    case "presentation":
      doctype = "presentation";
      break;
    case "pub":
      doctype = "pub";
      break;
    case "rtf":
      doctype = "rtf";
      break;
    case "spreadsheet":
      doctype = "spreadsheet";
      break;
    case "txt":
      doctype = "txt";
      break;
    case "vector":
      doctype = "vector";
      break;
    case "vsdx":
      doctype = "vsdx";
      break;
    case "vssx":
      doctype = "vssx";
      break;
    case "vstx":
      doctype = "vstx";
      break;
    case "xlsx":
      doctype = "xlsx";
      break;
    case "xltx":
      doctype = "xltx";
      break;
    case "xsn":
      doctype = "xsn";
      break;
    case "png":
      doctype = "photo";
      break;
    case "jpeg":
      doctype = "photo";
      break;
    case "jpg":
      doctype = "photo";
      break;
    case "img":
      doctype = "photo";
      break;
    default:
      doctype = "txt";
    
  }



  const url = `https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/assets/item-types/16/${doctype}.svg`;
  return url;
};


export default class SupportingDocumentsUploadFileComponent extends React.Component<
  IUploadFileProps,
  IUploadFileState
> {
  private fileInputRef: React.RefObject<HTMLInputElement>;

  public constructor(props: IUploadFileProps) {
    super(props);
    this.state = {
      selectedFiles: [],
      cummError: null,
      errorOfFile: null,
    };
    this.fileInputRef = React.createRef<HTMLInputElement>();
  }

  public componentDidMount(): void {
    this.validateFiles(this.props.data);
  }

  public componentDidUpdate(prevProps: IUploadFileProps): void {
    if (prevProps.data !== this.props.data) {
      this.validateFiles(this.props.data);
    }
  }

  private isFileNameValid(name: string): boolean {
    const regex = /^[a-zA-Z0-9._ -]+$/;
    return regex.test(name);
  }

  private validateFiles(files: File[]): void {
    const { maxFileSizeMB } = this.props;
    const maxFileSizeBytes = maxFileSizeMB * 1024 * 1024;

    let validFiles: IFileWithError[] = [];
    let currentTotalSize = 0;
    let cumulativeError = null;

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      let error: string | null = null;

      const allowedFileTypes = [".pdf", ".doc", ".docx", ".xlsx"];
      if (
        !allowedFileTypes.includes(
          file.name.substring(file.name.lastIndexOf("."))
        )
      ) {
        error = "File type is not allowed";
      } else if (file.size > maxFileSizeBytes) {
        error = `File size should not exceed more ${maxFileSizeMB}MB`;
      } else if (!this.isFileNameValid(file.name)) {
        error = "File name should not contain special characters";
      }

    

      if (
       
        currentTotalSize + file.size >
        maxFileSizeBytes
      ) {
        cumulativeError =
          "Cumulative size of all the supporting documents should not exceed 25 MB.";
      } else {
        cumulativeError = null;
      }

      currentTotalSize += file.size;
      validFiles.push({
        id: `${file.name}-${i}`,
        file,
        error,
        cumulativeError,
      });
     
      const filterNullerrorInvalidFiles = validFiles.filter((each: any) => {
        return each.error !== null;
      });
     
      this.props.errorData([filterNullerrorInvalidFiles, this.props.typeOfDoc]);
      this.props.cummulativeError(cumulativeError);
      this.setState({ errorOfFile: error, cummError: cumulativeError });
    }

    this.setState({ selectedFiles: validFiles, cummError: cumulativeError });
  }

  private handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files) {
      const files = Array.from(e.target.files);

     
      const hasAdditionalArray =
        this.props.addtionalData && this.props.addtionalData.length > 0;

      const newFiles = files.filter((file) => {
        const isDuplicateInSelectedFiles = this.state.selectedFiles.some(
          (selectedFile) => selectedFile.file.name === file.name
        );

        const isDuplicateInAnotherArray = hasAdditionalArray
          ? this.props.addtionalData.some(
              (anotherFile) => anotherFile.name === file.name
            )
          : false;

        
        return !isDuplicateInSelectedFiles && !isDuplicateInAnotherArray;
      });
      const filePromises = newFiles.map((file) =>
        this.convertToFileArrayBuffer(file)
      );

      Promise.all(filePromises)
        .then((fileBuffers) => {
          const filesWithBuffers = fileBuffers.map((buffer, index) => ({
            id: `${files[index].name}-${index}`,
            file: files[index],
            buffer: buffer,
            error: null,
            cumulativeError: null,
          }));

          const updatedFiles = this.props.multiple
            ? [...this.state.selectedFiles, ...filesWithBuffers]
            : filesWithBuffers;

            this.setState((prevState) => {
              const updatedFiles = this.props.multiple
                ? [...prevState.selectedFiles, ...filesWithBuffers]
                : filesWithBuffers;
              
              return { selectedFiles: updatedFiles };
            }, () => {
              this.validateFiles(this.state.selectedFiles.map((f) => f.file));
            });

          this.props.onChange(
            updatedFiles.map((f) => f.file),
            this.props.typeOfDoc
          );

          if (this.fileInputRef.current) {
            this.fileInputRef.current.value = "";
          }
        })
        .catch((error) => {
          console.error("Error converting files to ArrayBuffer", error);
        });
    }
  };

  
  private convertToFileArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        if (reader.result instanceof ArrayBuffer) {
          resolve(reader.result);
        } else {
          reject("FileReader result is not an ArrayBuffer");
        }
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  }

  

  private handleDeleteFile = (fileId: string): void => {
    const updatedFiles = this.state.selectedFiles.filter(
      (fileWithError) => fileWithError.id !== fileId
    );


    const { maxFileSizeMB } = this.props;
    const maxFileSizeBytes = maxFileSizeMB * 1024 * 1024;

    let cumulativeError = null;
    let currentTotalSize = 0;
    

    for (let i = 0; i < updatedFiles.length; i++) {
      const file = updatedFiles[i];
      

      if (
       
        currentTotalSize + file.file.size >
        maxFileSizeBytes
      ) {
        cumulativeError =
          "Cumulative size of all the supporting documents should not exceed 25 MB.";
      } else {
        cumulativeError = null;
      }
    }

    
    this.props.errorData([updatedFiles, this.props.typeOfDoc]);
    this.props.cummulativeError(cumulativeError);

    this.setState({ selectedFiles: updatedFiles }, () => {
      this.validateFiles(updatedFiles.map((f) => f.file));
    });

  

    this.props.onChange(
      updatedFiles.map((f) => f.file),
      this.props.typeOfDoc
    );
  };

  public render(): React.ReactElement<IUploadFileProps> {
    const { accept, typeOfDoc, multiple } = this.props;
    const { selectedFiles, cummError } = this.state;


    return (
      <ul className={`${styles.fileAttachementsUl}`}>
        <li className={`${styles.basicLi} ${styles.inputField}`}>
          <div style={{ padding: "8px" }}>
            <div>
              <button
                type="button"
                onClick={() => {
                  if (this.fileInputRef.current) {
                    this.fileInputRef.current.click();
                  }
                }}
              >
                Choose File
              </button>

              <input
                type="file"
                ref={this.fileInputRef}
                onChange={this.handleFileChange}
                accept={accept}
                multiple={multiple}
                style={{ display: "none" }}
              />
            </div>

            {typeOfDoc === "supportingDocument" &&
              cummError &&
              cummError.trim() !== "" && (
                <span
                  style={{
                    color: "red",
                    fontSize: "10px",
                   
                    margin: "0px",
                  }}
                >
                  {cummError}
                </span>
              )}
          </div>
        </li>

        {selectedFiles.length > 0 &&
          selectedFiles.map(({ id, file, error }) => {
         
            return (
              <li
                key={id}
                className={`${styles.basicLi} ${styles.attachementli}`}
              >
                <div className={`${styles.fileIconAndNameWithErrorContainer}`}>
                  <img
                  
                    src={_randomFileIcon(file.name)}
                    width={32}
                    height={32}
                  />
                 
                  <span className={`${styles.fileNameAndErrorContainer} `}>
                    <span
                      style={{
                        paddingBottom: "0px",
                        marginBottom: "0px",
                        paddingLeft: "4px",
                        whiteSpace: "nowrap",
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        display: "inline-block",
                       
                      }}
                    >
                    
                      {(file.name)}
                    </span>
                    {error && (
                      <span
                        style={{
                          color: "red",
                          fontSize: "10px",
                          paddingLeft: "4px",
                          margin: "0px",
                        }}
                      >
                        {error}
                      </span>
                    )}
                  </span>
                </div>

                <IconButton
                  iconProps={{ iconName: "Cancel" }}
                  title="Delete File"
                  ariaLabel="Delete File"
                  onClick={() => this.handleDeleteFile(id)}
                />
              </li>
            );
          })}
      </ul>
    );
  }
}
