import { useCallback, useEffect, useState } from "react";

export default function usePlans(fetchPlans) {
  const [plans, setPlans] = useState([]);
  const [status, setStatus] = useState("loading");
  const [error, setError] = useState(null);
  const [attempt, setAttempt] = useState(0);

  useEffect(() => {
    // FIX: every run owns a cancellation flag. It is checked before touching state, so a late
    // answer from a superseded request (reload) or an unmounted component cannot win the race.
    let cancelled = false;

    setPlans([]);
    setStatus("loading");
    setError(null);

    fetchPlans().then(
      (data) => {
        if (cancelled) return;
        setPlans(data);
        setStatus("ready");
        setError(null);
      },
      (err) => {
        // FIX: a rejected request now resolves the error state. The previous version only wrote
        // to the console and left the screen loading forever.
        if (cancelled) return;
        setPlans([]);
        setError(err);
        setStatus("error");
      }
    );

    // FIX: cleanup for the request (and for unmount), not just an empty function.
    return () => {
      cancelled = true;
    };
  }, [fetchPlans, attempt]);

  const reload = useCallback(() => setAttempt((n) => n + 1), []);

  return { status, plans, error, reload };
}
