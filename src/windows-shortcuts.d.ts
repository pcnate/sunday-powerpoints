declare module 'windows-shortcuts' {
  export const NORMAL: number;
  export const MAX: number;
  export const MIN: number;

  export interface ShortcutOptions {
    target: string;
    args?: string;
    cwd?: string;
    runStyle?: number;
    desc?: string;
    icon?: string;
    iconIndex?: number;
    hotkey?: string;
    windowStyle?: number;
    workingDir?: string;
    [key: string]: any;
  }

  export function create(
    lnkPath: string,
    options: ShortcutOptions,
    callback: (err: Error | null) => void
  ): void;

  export function query(
    lnkPath: string,
    callback: (err: Error | null, options?: ShortcutOptions) => void
  ): void;

  export function edit(
    lnkPath: string,
    options: ShortcutOptions,
    callback: (err: Error | null) => void
  ): void;
}
