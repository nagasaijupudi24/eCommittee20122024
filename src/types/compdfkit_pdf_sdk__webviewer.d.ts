declare module '@compdfkit_pdf_sdk/webviewer' {
  interface ComPDFKitViewerInstance {
    docViewer: {
      addEvent: (event: string, callback: () => void) => void;
    };
  }

  const ComPDFKitViewer: {
    init: (options: { path: string; pdfUrl: string; license: string }, container: HTMLElement | null) => Promise<ComPDFKitViewerInstance>;
  };

  export default ComPDFKitViewer;
}
