import { useCallback, useEffect, useRef, useState } from "react";

const IDLE = { status: "idle", plan: null, result: null, error: null };

export default function useAgentRun(runAgent) {
  const [run, setRun] = useState(IDLE);
  const latest = useRef(IDLE);
  const controller = useRef(null);

  const update = useCallback((next) => {
    latest.current = next;
    setRun(next);
  }, []);

  const start = useCallback(
    (plan) => {
      const { status } = latest.current;
      if (status === "running" || status === "awaiting_approval") return;
      // FIX: the missing runAgent is a guard inside start(), not an early return before the
      // hooks. The early return changed the number of hooks between renders.
      if (!runAgent) return;

      const current = new AbortController();
      controller.current = current;
      update({ status: "running", plan, result: null, error: null });

      runAgent(plan, { signal: current.signal }).then(
        (result) => {
          if (current.signal.aborted) return;
          // FIX: an R2 plan never goes from running to done. It waits for approve().
          if (plan.risk === "R2") {
            update({ status: "awaiting_approval", plan, result, error: null });
          } else {
            update({ status: "done", plan, result, error: null });
          }
        },
        (error) => {
          if (current.signal.aborted) return;
          update({ status: "error", plan, result: null, error });
        }
      );
    },
    [runAgent, update]
  );

  const approve = useCallback(() => {
    if (latest.current.status !== "awaiting_approval") return;
    update({ ...latest.current, status: "done" });
  }, [update]);

  const reject = useCallback(() => {
    const { status } = latest.current;
    if (status !== "running" && status !== "awaiting_approval") return;
    controller.current?.abort();
    controller.current = null;
    update(IDLE);
  }, [update]);

  // FIX: abort the run in flight when the component unmounts. The previous version had no
  // cleanup, so a late answer could still call setRun on an unmounted component.
  useEffect(() => {
    return () => {
      controller.current?.abort();
      controller.current = null;
    };
  }, []);

  return { ...run, start, approve, reject };
}
