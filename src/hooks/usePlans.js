import { useCallback, useEffect, useState } from "react";

const INITIAL = { status: "loading", plans: [], error: null };

/**
 * Loads the agent plans and exposes the three states of a request: loading, ready and error.
 *
 * `fetchPlans` must be a stable function (module level or wrapped in useCallback); when it
 * changes the hook asks again. `reload()` repeats the request and returns to loading.
 *
 * Every effect run owns a `cancelled` flag. The cleanup sets it to true, so the answer of a
 * superseded request (a reload or an unmount) can never overwrite the newest one: that is the
 * request race the tests call "la respuesta vieja".
 */
export default function usePlans(fetchPlans) {
  const [state, setState] = useState(INITIAL);
  const [attempt, setAttempt] = useState(0);

  useEffect(() => {
    let cancelled = false;
    setState(INITIAL);

    fetchPlans().then(
      (plans) => {
        if (!cancelled) setState({ status: "ready", plans, error: null });
      },
      (error) => {
        if (!cancelled) setState({ status: "error", plans: [], error });
      }
    );

    return () => {
      cancelled = true;
    };
  }, [fetchPlans, attempt]);

  const reload = useCallback(() => {
    setAttempt((current) => current + 1);
  }, []);

  return { ...state, reload };
}
