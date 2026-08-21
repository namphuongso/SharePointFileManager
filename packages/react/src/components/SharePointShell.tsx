import { useMemo, useState, type ReactNode } from "react";
import {
  Avatar,
  Button,
  Drawer,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  Input,
  Text,
  tokens,
} from "@fluentui/react-components";
import {
  ChevronLeftRegular,
  ChevronRightRegular,
  NavigationRegular,
  SearchRegular,
} from "@fluentui/react-icons";

export interface SharePointShellHeaderConfig {
  appName: string;
  logo?: ReactNode;
  showSearch?: boolean;
  showUserMenu?: boolean;
}

export interface SharePointShellNavigationItem {
  key: string;
  label: string;
  icon?: ReactNode;
  badge?: string | number;
  disabled?: boolean;
  children?: SharePointShellNavigationItem[];
  customRender?: (item: SharePointShellNavigationItem, active: boolean) => ReactNode;
}

interface SharePointShellProps {
  header: SharePointShellHeaderConfig;
  navigationItems: SharePointShellNavigationItem[];
  activeNavigationKey?: string;
  collapsibleNavigation?: boolean;
  showHeader?: boolean;
  showNavigation?: boolean;
  onNavigationChange?: (key: string) => void;
  onHeaderSearch?: (query: string) => void;
  children: ReactNode;
}

export function SharePointShell({
  header,
  navigationItems,
  activeNavigationKey,
  collapsibleNavigation = true,
  showHeader = true,
  showNavigation = true,
  onNavigationChange,
  onHeaderSearch,
  children,
}: SharePointShellProps) {
  const [collapsed, setCollapsed] = useState(false);
  const [drawerOpen, setDrawerOpen] = useState(false);

  const headerSearchEnabled = header.showSearch !== false;
  const userMenuEnabled = header.showUserMenu !== false;

  const navigation = useMemo(
    () => (
      <nav className={`spm-side-nav ${collapsed ? "collapsed" : ""}`} aria-label="Library navigation">
        <div className="spm-side-nav-list">
          {navigationItems.map((item) => {
            const active = item.key === activeNavigationKey;
            if (item.customRender) {
              return <div key={item.key}>{item.customRender(item, active)}</div>;
            }
            return (
              <button
                key={item.key}
                type="button"
                disabled={item.disabled}
                className={`spm-side-nav-item ${active ? "active" : ""}`}
                onClick={() => onNavigationChange?.(item.key)}
                title={collapsed ? item.label : undefined}
              >
                <span className="spm-side-nav-icon">{item.icon}</span>
                {!collapsed ? <span className="spm-side-nav-label">{item.label}</span> : null}
                {!collapsed && item.badge !== undefined ? (
                  <span className="spm-side-nav-badge">{item.badge}</span>
                ) : null}
              </button>
            );
          })}
        </div>
      </nav>
    ),
    [activeNavigationKey, collapsed, navigationItems, onNavigationChange],
  );

  return (
    <div className="spm-shell">
      {showHeader ? (
        <header className="spm-m365-header" style={{ borderColor: tokens.colorNeutralStroke2 }}>
        <div className="spm-m365-header-left">
          <Button
            appearance="subtle"
            icon={<NavigationRegular />}
            aria-label="Open navigation"
            className="spm-mobile-nav-trigger"
            onClick={() => setDrawerOpen(true)}
          />
          <div className="spm-app-brand">
            {header.logo ? <span className="spm-app-logo">{header.logo}</span> : null}
            <Text weight="semibold">{header.appName}</Text>
          </div>
        </div>
        <div className="spm-m365-header-right">
          {headerSearchEnabled ? (
            <Input
              className="spm-m365-search"
              contentBefore={<SearchRegular />}
              placeholder="Search"
              onChange={(_, data) => onHeaderSearch?.(data.value)}
            />
          ) : null}
          {userMenuEnabled ? <Avatar name={header.appName} size={28} /> : null}
        </div>
        </header>
      ) : null}

      <div className="spm-shell-main">
        {showNavigation ? navigation : null}
        {showNavigation && collapsibleNavigation ? (
          <Button
            appearance="subtle"
            className="spm-side-nav-toggle"
            icon={collapsed ? <ChevronRightRegular /> : <ChevronLeftRegular />}
            onClick={() => setCollapsed((current) => !current)}
            aria-label={collapsed ? "Expand navigation" : "Collapse navigation"}
          />
        ) : null}
        <section className="spm-shell-content">{children}</section>
      </div>

      <Drawer
        modalType="modal"
        open={drawerOpen}
        onOpenChange={(_, data) => setDrawerOpen(data.open)}
        position="start"
      >
        <DrawerHeader>
          <DrawerHeaderTitle>{header.appName}</DrawerHeaderTitle>
        </DrawerHeader>
        <DrawerBody>{navigation}</DrawerBody>
      </Drawer>
    </div>
  );
}
