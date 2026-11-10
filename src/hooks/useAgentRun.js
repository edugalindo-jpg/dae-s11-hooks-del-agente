import { useCallback, useEffect, useRef, useState } from "react";

const IDLE = { status: "idle", plan: null, result: null, error: null };

/**
 * State machine for one agent run:
 *
 *   idle -> running -> done                       (R1)
 *   idle -> running -> awaiting_approval -> done  (R2, only after approve())
 *   idle -> running -> error
 *
 * `latest` mirrors the state so start/approve/reject can read it without waiting for a render;
 * `controller` owns the AbortSignal of the current run. Both are refs, so the hook order never
 * changes even when `runAgent` is still undefined.
 */
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
      if (!runAgent) return;

      const current = new AbortController();
      controller.current = current;
      update({ status: "running", plan, result: null, error: null });

      runAgent(plan, { signal: current.signal }).then(
        (result) => {
          if (current.signal.aborted) return;
          // Safety rule: a destructive plan (R2) never goes straight to done.
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

  // Unmount: abort the run in flight. Its promise checks `signal.aborted` before touching state.
  useEffect(() => {
    return () => {
      controller.current?.abort();
      controller.current = null;
    };
  }, []);

  return { ...run, start, approve, reject };
}
