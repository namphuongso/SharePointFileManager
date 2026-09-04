import { useEffect, useId, useMemo, useRef, type ReactNode } from "react";
import {
  Button as FluentButton,
  Toaster,
  useToastController,
  type ToastId,
  type ToastIntent,
  type ToastPosition,
} from "@fluentui/react-components";
import {
  CheckmarkCircle24Filled,
  DismissRegular,
  ErrorCircle24Filled,
  Info24Filled,
} from "@fluentui/react-icons";
import { useToastStyles } from "./toastStyles";
import type { NotifyApi } from "./types";

/** Auto-dismiss mặc định — thành công 3s, lỗi 6s; tiến trình không tự tắt. */
const TIMEOUT_SUCCESS = 3000;
const TIMEOUT_INFO = 3000;
const TIMEOUT_ERROR = 6000;
const TIMEOUT_STICKY = -1;

/** Số toast tối đa hiển thị đồng thời — bảng bên dưới không bị đầy. */
const TOAST_LIMIT = 5;

/** Góc dưới-phải, cách mép 16px — giống Teams/Outlook. */
const TOAST_POSITION = "bottom-end" as ToastPosition;
const TOAST_OFFSET = { horizontal: 16, vertical: 16 };

interface ToastCardProps {
  intent: ToastIntent;
  title: string;
  subtitle?: string;
  icon: ReactNode;
  toastId: ToastId;
  onClose: (id: ToastId) => void;
}

function ToastCard({ intent, title, subtitle, icon, toastId, onClose }: ToastCardProps) {
  const styles = useToastStyles();
  const mediaClass =
    intent === "success"
      ? styles.mediaSuccess
      : intent === "error"
        ? styles.mediaError
        : styles.mediaInfo;
  return (
    <div className={styles.root} role={intent === "error" ? "alert" : "status"}>
      <span className={`${styles.media} ${mediaClass}`} aria-hidden>
        {icon}
      </span>
      <div className={styles.body}>
        <span className={styles.title}>{title}</span>
        {subtitle ? <span className={styles.subtitle}>{subtitle}</span> : null}
      </div>
      <FluentButton
        appearance="subtle"
        shape="circular"
        size="medium"
        className={styles.close}
        icon={<DismissRegular fontSize={18} />}
        aria-label="Đóng thông báo"
        onClick={() => onClose(toastId)}
      />
    </div>
  );
}

function makeToastId(toasterId: string): ToastId {
  return `${toasterId}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

export interface ToastViewportProps {
  /** Nhận API dispatch khi controller sẵn sàng. Provider dùng để đưa vào context. */
  onReady: (api: NotifyApi) => void;
}

/**
 * Mount một `Toaster` ở góc dưới-phải + trả về API `NotifyApi` qua callback.
 * Phải nằm dưới `FluentProvider` để dùng token màu.
 */
export function ToastViewport({ onReady }: ToastViewportProps): ReactNode {
  const toasterId = useId();
  const { dispatchToast, dismissToast, updateToast } = useToastController(toasterId);
  const dismissRef = useRef(dismissToast);
  const updateRef = useRef(updateToast);
  dismissRef.current = dismissToast;
  updateRef.current = updateToast;

  const api = useMemo<NotifyApi>(
    () => ({
      success: (title, subtitle) => {
        const id = makeToastId(toasterId);
        const icon = <CheckmarkCircle24Filled />;
        dispatchToast(
          <ToastCard
            intent="success"
            title={title}
            subtitle={subtitle}
            icon={icon}
            toastId={id}
            onClose={(tid) => dismissRef.current(tid)}
          />,
          {
            toastId: id,
            intent: "success",
            position: TOAST_POSITION,
            timeout: TIMEOUT_SUCCESS,
            pauseOnHover: true,
          },
        );
      },
      info: (title, subtitle) => {
        const id = makeToastId(toasterId);
        const icon = <Info24Filled />;
        dispatchToast(
          <ToastCard
            intent="info"
            title={title}
            subtitle={subtitle}
            icon={icon}
            toastId={id}
            onClose={(tid) => dismissRef.current(tid)}
          />,
          {
            toastId: id,
            intent: "info",
            position: TOAST_POSITION,
            timeout: TIMEOUT_INFO,
            pauseOnHover: true,
          },
        );
        return id;
      },
      progress: (title, subtitle) => {
        const id = makeToastId(toasterId);
        const icon = <Info24Filled />;
        dispatchToast(
          <ToastCard
            intent="info"
            title={title}
            subtitle={subtitle}
            icon={icon}
            toastId={id}
            onClose={(tid) => dismissRef.current(tid)}
          />,
          {
            toastId: id,
            intent: "info",
            position: TOAST_POSITION,
            timeout: TIMEOUT_STICKY,
            pauseOnHover: true,
          },
        );
        return id;
      },
      error: (title, subtitle) => {
        const id = makeToastId(toasterId);
        const icon = <ErrorCircle24Filled />;
        dispatchToast(
          <ToastCard
            intent="error"
            title={title}
            subtitle={subtitle}
            icon={icon}
            toastId={id}
            onClose={(tid) => dismissRef.current(tid)}
          />,
          {
            toastId: id,
            intent: "error",
            position: TOAST_POSITION,
            timeout: TIMEOUT_ERROR,
            pauseOnHover: true,
            politeness: "assertive",
          },
        );
      },
      update: (id, options) => {
        const icon =
          options.intent === "success" ? (
            <CheckmarkCircle24Filled />
          ) : options.intent === "error" ? (
            <ErrorCircle24Filled />
          ) : (
            <Info24Filled />
          );
        // info = vẫn đang xử lý → sticky; success/error → auto-dismiss.
        const timeout =
          options.intent === "error"
            ? TIMEOUT_ERROR
            : options.intent === "success"
              ? TIMEOUT_SUCCESS
              : TIMEOUT_STICKY;
        updateRef.current({
          toastId: id,
          intent: options.intent,
          content: (
            <ToastCard
              intent={options.intent}
              title={options.title}
              subtitle={options.subtitle}
              icon={icon}
              toastId={id}
              onClose={(tid) => dismissRef.current(tid)}
            />
          ),
          position: TOAST_POSITION,
          timeout,
          pauseOnHover: true,
          politeness: options.intent === "error" ? "assertive" : undefined,
        });
      },
      dismiss: (id) => {
        dismissRef.current(id);
      },
    }),
    [dispatchToast, toasterId],
  );

  useEffect(() => {
    onReady(api);
  }, [api, onReady]);

  return (
    <Toaster
      toasterId={toasterId}
      position={TOAST_POSITION}
      limit={TOAST_LIMIT}
      offset={TOAST_OFFSET}
    />
  );
}
