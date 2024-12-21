/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import { useEffect, useRef, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.entry";
import styles from "./PdfViewer.module.scss";
import { Icon } from "@fluentui/react";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

const PDFViewer: React.FC<{ pdfPath: string; noteNumber: string }> = (props) => {
  const { pdfPath, noteNumber } = props;

  const pdfViewerRef = useRef<HTMLDivElement>(null);
  const [pdfDocument, setPdfDocument] = useState<pdfjsLib.PDFDocumentProxy | null>(null);
  const [numPages, setNumPages] = useState<number>(0);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [zoomLevel, setZoomLevel] = useState<number>(1.1); 
  const [renderedPages, setRenderedPages] = useState<Map<number, string>>(new Map());
 
  const zoomLevels = [
    { value: "1.3", label: "Actual Width" },
    { value: "1.2", label: "Fit to Width" },
    { value: "1.1", label: "Fit to Page" },
    { value: "0.5", label: "50%" },
    { value: "0.75", label: "75%" },
    { value: "1", label: "100%" },
    { value: "1.25", label: "125%" },
    { value: "1.5", label: "150%" },
    { value: "2", label: "200%" },
    { value: "3", label: "300%" },
    { value: "4", label: "400%" },
  ];

  
  useEffect(() => {
    const loadPdf = async () => {
      const loadingTask = pdfjsLib.getDocument(pdfPath);
      const pdf = await loadingTask.promise;
      setPdfDocument(pdf);
      setNumPages(pdf.numPages);
    };
    loadPdf().catch(console.error);
  }, [pdfPath]);

  
  const renderPage = async (pageNum: number) => {
    if (!pdfDocument || renderedPages.has(pageNum)) return;
  
    const page = await pdfDocument.getPage(pageNum);
    const viewport = page.getViewport({ scale: zoomLevel });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
  
    if (context) {
      canvas.height = viewport.height;
      canvas.width = viewport.width;
      await page.render({ canvasContext: context, viewport }).promise;
  
    
      setRenderedPages((prev) => new Map(prev).set(pageNum, canvas.toDataURL()));
    }
  };
  

  const renderAllPages = async () => {
    if (!pdfDocument) return;
  
   
  
    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      await renderPage(pageNum);
    }
  };
  
 
  useEffect(() => {
    renderAllPages().catch(console.error);
  }, [zoomLevel, pdfDocument, numPages]);
  


  const handleScroll = () => {
    if (!pdfViewerRef.current) return;

    const viewer = pdfViewerRef.current;
    const viewerTop = viewer.scrollTop;
    const viewerBottom = viewerTop + viewer.clientHeight;

    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      const pageHeight = viewer.scrollHeight / numPages;
      const pageTop = (pageNum - 1) * pageHeight;
      const pageBottom = pageTop + pageHeight;

      if (
        (pageTop >= viewerTop && pageTop <= viewerBottom) ||
        (pageBottom >= viewerTop && pageBottom <= viewerBottom)
      ) {
        renderPage(pageNum).catch(console.error);
        setCurrentPage(pageNum);
      }
    }
  };

  useEffect(() => {
    const pdfViewer = pdfViewerRef.current;
    pdfViewer?.addEventListener("scroll", handleScroll);
    return () => {
      pdfViewer?.removeEventListener("scroll", handleScroll);
    };
  }, [numPages, zoomLevel]);


  const handleZoomChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    setZoomLevel(parseFloat(event.target.value));
    setRenderedPages(new Map());
  };


  const resetAndRenderPages = async (newZoomLevel: number) => {
    setRenderedPages(new Map());
    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      await renderPage(pageNum); 
    }
    setCurrentPage(currentPage); 
  };

  const handleZoomIn = () => {
    const currentZoomIndex = zoomLevels.findIndex(level => parseFloat(level.value) === zoomLevel);
    if (currentZoomIndex < zoomLevels.length - 1) {
      const newZoomLevel = parseFloat(zoomLevels[currentZoomIndex + 1].value);
      setZoomLevel(newZoomLevel);
      resetAndRenderPages(newZoomLevel);
    }
    setCurrentPage(currentPage); 
  };
  
  const handleZoomOut = () => {
    const currentZoomIndex = zoomLevels.findIndex(level => parseFloat(level.value) === zoomLevel);
    if (currentZoomIndex > 0) {
      const newZoomLevel = parseFloat(zoomLevels[currentZoomIndex - 1].value);
      setZoomLevel(newZoomLevel);
      resetAndRenderPages(newZoomLevel);
    }
    setCurrentPage(currentPage); 
  };

  const handleNextPage = () => {
    const nextPage = Math.min(currentPage + 1, numPages);
    setCurrentPage(nextPage);
    scrollToPage(nextPage).catch(console.error);
  };
  
  const handlePreviousPage = () => {
    const prevPage = Math.max(currentPage - 1, 1);
    setCurrentPage(prevPage);
    scrollToPage(prevPage).catch(console.error);
  };
  

  const scrollToPage = async (pageNum: number) => {
    if (!pdfViewerRef.current || !pdfDocument) return;
  
    const page = await pdfDocument.getPage(pageNum);
    const viewport = page.getViewport({ scale: zoomLevel });
  
   
    const pageHeight = viewport.height;
    const targetScrollTop = pageHeight * (pageNum - 1);
  
    pdfViewerRef.current.scrollTo({
      top: targetScrollTop,
      behavior: "smooth",
    });
  };
  

  const handleSave = () => {
    const a = document.createElement("a");
    a.href = pdfPath;
    a.download = `${noteNumber}.pdf`;
    a.click();
  };

  const handlePrint = async () => {
    console.log('Print triggered');
    
   
    if (!pdfDocument) return;

    const printWindow = window.open('', '_blank');
    if (!printWindow) return;

    const printContent = document.createElement('div');
    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      await renderPage(pageNum);
      const imgSrc = renderedPages.get(pageNum);
      if (imgSrc) {
        const img = new Image();
        img.src = imgSrc;
        img.style.width = '100%';
        printContent.appendChild(img);
        printContent.appendChild(document.createElement('br'));
      }
    }

    printWindow.document.write('<html><head><title>Print PDF</title></head><body></body></html>');
    printWindow.document.body.appendChild(printContent);
    printWindow.document.close();

    printWindow.onload = () => {
      printWindow.print();
      printWindow.close();
    };


};

  

  return (
    
    <div>
      <div className={styles.toolbar}>
        <div id={styles.toolbarContainer}>
          <div id={styles.toolbarViewer}>
            <div id={styles.toolbarViewerLeft}>
              <button
                className={styles.toolbarButton}
                title="Previous Page"
                onClick={handlePreviousPage}
                disabled={currentPage <= 1}
              >
                <Icon iconName="ChevronUp" />
              </button>
              <button
                className={styles.toolbarButton}
                title="Next Page"
                onClick={handleNextPage}
                disabled={currentPage >= numPages}
              >
                <Icon iconName="ChevronDown" />
              </button>
              <span className={styles.toolbarLabel}>
                {currentPage} / {numPages}
              </span>
            </div>
            <div id={styles.toolbarViewerMiddle}>
              <button className={styles.toolbarButton} title="Zoom Out" onClick={handleZoomOut}>
                <Icon iconName="Remove" />
              </button>
              <button className={styles.toolbarButton} title="Zoom In" onClick={handleZoomIn}>
                <Icon iconName="Add" />
              </button>
              <div className={styles.dropdownToolbarButton}>
                <span id="scaleSelectContainer" className={styles.dropdownToolbarButton}>
                  <select
                    id="scaleSelect"
                    title="Zoom"
                    value={zoomLevel.toString()}
                    onChange={handleZoomChange}
                  >
                    {zoomLevels.map((zoom) => (
                      <option key={zoom.value} value={zoom.value}>
                        {zoom.label}
                      </option>
                    ))}
                  </select>
                </span>
              </div>
            </div>
            <div id={styles.toolbarViewerRight}>
              <button className={styles.toolbarButton} title="Print" onClick={handlePrint}>
                <Icon iconName="Print" />
              </button>
              <button className={styles.toolbarButton} title="Download" onClick={handleSave}>
                <Icon iconName="Save" />
              </button>
            </div>
          </div>
        </div>
      </div>
      <div className={styles.pdfviewer} ref={pdfViewerRef}>
        {Array.from(renderedPages.entries()).map(([pageNum, imgSrc]) => (
          <img key={pageNum} src={imgSrc} alt={`Page ${pageNum}`} />
        ))}
      </div>
    </div>
  );
};

export default PDFViewer;
