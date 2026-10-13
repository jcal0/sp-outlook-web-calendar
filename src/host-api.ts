// The parts of Super Productivity's plugin API this plugin uses. The npm package
// @super-productivity/plugin-api (1.0.1) predates issue providers, so these are taken
// from packages/plugin-api/src in the Super Productivity repo (MIT). Keep them to
// what the plugin calls.

export interface PluginSearchResult {
  id: string;
  title: string;
  /** Event start (ms); required for calendar display */
  start?: number;
  /** Timed events: the task created from it is scheduled at this time */
  dueWithTime?: number;
  /** Event duration (ms) */
  duration?: number;
  isAllDay?: boolean;
  description?: string;
  [key: string]: unknown;
}

export interface PluginIssue {
  id: string;
  title: string;
  body?: string;
  state?: string;
  [key: string]: unknown;
}

export interface PluginIssueField {
  field: string;
  label: string;
  type?: 'text' | 'markdown' | 'link' | 'date' | 'list';
  hideEmpty?: boolean;
}

export interface IssueProviderPluginDefinition {
  configFields: unknown[];
  getHeaders(config: Record<string, unknown>): Record<string, string>;
  searchIssues(searchTerm: string): Promise<PluginSearchResult[]>;
  getById(issueId: string): Promise<PluginIssue>;
  getIssueLink(issueId: string): string;
  getNewIssuesForBacklog(): Promise<PluginSearchResult[]>;
  issueDisplay: PluginIssueField[];
}

export interface DialogButtonCfg {
  label: string;
  icon?: string;
  onClick?: () => void | Promise<void>;
  color?: 'primary' | 'warn';
  raised?: boolean;
}

export interface PluginNodeScriptResult {
  success: boolean;
  result?: unknown;
  error?: string;
}

export interface HostPluginAPI {
  registerIssueProvider(definition: IssueProviderPluginDefinition): void;
  registerHeaderButton(cfg: { label: string; icon?: string; onClick: () => void }): void;
  registerMenuEntry(cfg: { label: string; icon?: string; onClick: () => void }): void;
  registerShortcut(cfg: { id?: string; label: string; onExec: () => void }): void;
  showSnack(cfg: {
    msg: string;
    type?: 'SUCCESS' | 'ERROR' | 'INFO' | 'WARNING';
    ico?: string;
  }): void;
  openDialog(cfg: {
    title?: string;
    htmlContent?: string;
    buttons?: DialogButtonCfg[];
  }): Promise<string | undefined>;
  getConfig<T = Record<string, unknown>>(): Promise<T | null>;
  formatDate(
    date: Date | string | number,
    format: 'short' | 'medium' | 'long' | 'time' | 'datetime',
  ): string;
  /** Desktop app only (nodeExecution permission + user consent) */
  executeNodeScript?(request: {
    script: string;
    args?: unknown[];
    timeout?: number;
  }): Promise<PluginNodeScriptResult>;
  log: { info: (...args: unknown[]) => void };
}
