import { useEffect, useRef } from 'react';

// <comofficekit-viewer> as a React component: props become the element's attributes, its events become callbacks,
// and `file` (a File or Blob the user chose) is sent with open(). Any other prop goes to the element: style, className,
// zoom, credentials, toolbar.
export default function DocumentViewer({ src, documentId, file, onLoad, onError, onPageChange, ...rest }) {
  const element = useRef(null);
  // The latest callbacks, so that a new function on each render does not add listeners again.
  const callbacks = useRef({});
  callbacks.current = { onLoad, onError, onPageChange };

  useEffect(() => {
    const viewer = element.current;
    const load = event => callbacks.current.onLoad?.(event.detail);
    const error = event => callbacks.current.onError?.(event.detail);
    const pageChange = event => callbacks.current.onPageChange?.(event.detail.page);
    viewer.addEventListener('load', load);
    viewer.addEventListener('error', error);
    viewer.addEventListener('pagechange', pageChange);
    return () => {
      viewer.removeEventListener('load', load);
      viewer.removeEventListener('error', error);
      viewer.removeEventListener('pagechange', pageChange);
    };
  }, []);

  useEffect(() => {
    // The element says what went wrong in its toolbar and in its error event.
    if (file) element.current.open(file).catch(() => {});
  }, [file]);

  return <comofficekit-viewer ref={element} src={src} document={documentId} {...rest} />;
}
