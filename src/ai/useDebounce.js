import { useEffect, useState } from "react";

export default function useDebounce(value, delay) {
  const [debounced, setDebounced] = useState(value);

  useEffect(() => {
    const timer = setTimeout(() => {
      setDebounced(value);
    }, delay);

    // FIX: the timer is cancelled on every change and on unmount, not only on unmount.
    // The previous version kept a cleanup effect with [] and stored the timer in a ref, so the
    // timers of superseded values stayed alive and emitted intermediate values later.
    return () => clearTimeout(timer);
  }, [value, delay]);

  return debounced;
}
