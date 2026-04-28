import { useState, useEffect } from 'react';

// Using a module-level cache to track script loading status across component instances
const cachedScriptStatuses: { [key: string]: 'loading' | 'ready' | 'error' } = {};

const SCRIPT_LOAD_TIMEOUT = 10000; // 10 seconds

/**
 * A custom hook to dynamically load an external script and track its status.
 * It prevents re-adding the script if it's already in the DOM or being loaded.
 * Includes timeout handling for better error detection.
 * @param src The URL of the script to load.
 * @returns The loading status: 'loading', 'ready', or 'error'.
 */
export function useScript(src: string): 'loading' | 'ready' | 'error' {
  const [status, setStatus] = useState<'loading' | 'ready' | 'error'>(() => {
    // Return cached status if it exists, otherwise default to 'loading'
    return cachedScriptStatuses[src] || 'loading';
  });

  useEffect(() => {
    // If this script's status is already resolved, we don't need to do anything.
    const cachedStatus = cachedScriptStatuses[src];
    if (cachedStatus && (cachedStatus === 'ready' || cachedStatus === 'error')) {
      setStatus(cachedStatus);
      return;
    }

    // Find an existing script tag in the document
    let script = document.querySelector(`script[src="${src}"]`) as HTMLScriptElement;

    // If the script doesn't exist, create it
    if (!script) {
      script = document.createElement('script');
      script.src = src;
      script.async = true;
      document.body.appendChild(script);

      // Set its status to 'loading' in our cache
      cachedScriptStatuses[src] = 'loading';
    }

    // Set up timeout for script loading
    const timeoutId = setTimeout(() => {
      if (cachedScriptStatuses[src] === 'loading') {
        cachedScriptStatuses[src] = 'error';
        setStatus('error');
        console.error(`Script loading timeout: ${src}`);
      }
    }, SCRIPT_LOAD_TIMEOUT);

    // Event listener for successful script loading
    const handleLoad = () => {
      clearTimeout(timeoutId);
      cachedScriptStatuses[src] = 'ready';
      setStatus('ready');
    };

    // Event listener for script loading failure
    const handleError = () => {
      clearTimeout(timeoutId);
      cachedScriptStatuses[src] = 'error';
      setStatus('error');
      console.error(`Script loading failed: ${src}`);
    };

    script.addEventListener('load', handleLoad);
    script.addEventListener('error', handleError);

    // Cleanup function to remove event listeners when the component unmounts
    return () => {
      clearTimeout(timeoutId);
      if (script) {
        script.removeEventListener('load', handleLoad);
        script.removeEventListener('error', handleError);
      }
    };
  }, [src]); // Re-run effect only if the script 'src' changes

  return status;
}
