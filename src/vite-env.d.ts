/// <reference types="vite/client" />

declare module "ssf" {
  export function format(fmt: string | number, value: unknown, opts?: Record<string, unknown>): string;
  export function is_date(fmt: string): boolean;
}
