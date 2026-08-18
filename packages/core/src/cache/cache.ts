export interface CacheProvider {
  get<T>(key: string): Promise<T | undefined> | T | undefined;
  set<T>(key: string, value: T, ttlMs?: number): Promise<void> | void;
  delete(key: string): Promise<void> | void;
  clear(): Promise<void> | void;
  keys(): Promise<string[]> | string[];
}

interface Entry { value: unknown; expiresAt?: number }

export class MemoryCacheProvider implements CacheProvider {
  private readonly entries = new Map<string, Entry>();
  get<T>(key: string): T | undefined {
    const entry = this.entries.get(key);
    if (!entry) return undefined;
    if (entry.expiresAt !== undefined && entry.expiresAt <= Date.now()) {
      this.entries.delete(key);
      return undefined;
    }
    return entry.value as T;
  }
  set<T>(key: string, value: T, ttlMs?: number): void {
    this.entries.set(key, { value, expiresAt: ttlMs ? Date.now() + ttlMs : undefined });
  }
  delete(key: string): void { this.entries.delete(key); }
  clear(): void { this.entries.clear(); }
  keys(): string[] { return [...this.entries.keys()]; }
}
