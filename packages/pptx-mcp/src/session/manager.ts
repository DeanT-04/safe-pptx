import fs from 'node:fs';
import { PptxPackage } from '@safe-pptx/pptx-core';

export interface EditRecord {
  timestamp: string;
  tool: string;
  target: string;
  before: string;
  after: string;
}

export interface PptxSession {
  /** Absolute, policy-resolved file path (session key). */
  filePath: string;
  pkg: PptxPackage;
  aiAuthor: string;
  createdAt: string;
  lastAccess: number;
  audit: EditRecord[];
}

export const SESSION_TTL_MS = 60 * 60 * 1000; // 1 hour of inactivity, like safe-docx

export class SessionManager {
  private readonly sessions = new Map<string, PptxSession>();
  private readonly timer: NodeJS.Timeout;

  constructor(readonly aiAuthor = 'safe-pptx AI') {
    this.timer = setInterval(() => this.sweep(), 60_000);
    this.timer.unref();
  }

  stats(): { sessions: number; paths: string[] } {
    return { sessions: this.sessions.size, paths: [...this.sessions.keys()].sort() };
  }

  async getOrCreate(filePath: string): Promise<PptxSession> {
    this.sweep();
    const existing = this.sessions.get(filePath);
    if (existing) {
      existing.lastAccess = Date.now();
      return existing;
    }
    const buffer = fs.readFileSync(filePath);
    const pkg = await PptxPackage.load(buffer);
    const session: PptxSession = {
      filePath,
      pkg,
      aiAuthor: this.aiAuthor,
      createdAt: new Date().toISOString(),
      lastAccess: Date.now(),
      audit: [],
    };
    this.sessions.set(filePath, session);
    return session;
  }

  get(filePath: string): PptxSession | undefined {
    this.sweep();
    const session = this.sessions.get(filePath);
    if (session) session.lastAccess = Date.now();
    return session;
  }

  require(filePath: string): PptxSession {
    const session = this.get(filePath);
    if (!session) {
      throw new Error(
        `no open session for ${filePath}. Sessions open automatically on first use — ` +
          'if the file was never opened, any read/edit tool call will create one.',
      );
    }
    return session;
  }

  close(filePath: string): boolean {
    return this.sessions.delete(filePath);
  }

  closeAll(): number {
    const n = this.sessions.size;
    this.sessions.clear();
    return n;
  }

  sweep(): void {
    const cutoff = Date.now() - SESSION_TTL_MS;
    for (const [key, session] of this.sessions) {
      if (session.lastAccess < cutoff) this.sessions.delete(key);
    }
  }

  stop(): void {
    clearInterval(this.timer);
  }
}

export const sessionManager = new SessionManager();
