import { useEffect, useState } from "react";

/**
 * Returns the "quiet" value: the last `value` that stayed unchanged for `delay` ms.
 *
 * The returned value is state, initialised with `value` so the first render never waits.
 * Every change cancels the previous timer and starts a new one; the cleanup cancels the
 * timer when the effect repeats and when the component unmounts, so no intermediate value
 * is ever emitted and no timer survives the hook.
 */
export default function useDebounce(value, delay) {
  const [debounced, setDebounced] = useState(value);

  useEffect(() => {
    const timer = setTimeout(() => {
      setDebounced(value);
    }, delay);

    return () => clearTimeout(timer);
  }, [value, delay]);

  return debounced;
}
