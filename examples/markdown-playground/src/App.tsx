import { useCallback, useEffect, useMemo, useState } from 'react';
import { DocxEditor } from '@eigenpal/docx-editor-react';
import {
  toMarkdown,
  toMarkdownPaged,
  type MarkdownResult,
  type PagedMarkdownResult,
} from '@eigenpal/docx-editor-core/markdown';

type Method = 'continuous' | 'paged';
type Tracked = 'clean' | 'annotate';
type Comments = 'strip' | 'inline' | 'sidecar';
type Hyperlinks = 'inline' | 'reference';
type Annotations = 'html' | 'pandoc' | 'strip';
type Footnotes = 'inline' | 'end';
type HeaderFooter = 'strip' | 'first-page' | 'all';

interface Options {
  trackedChanges: Tracked;
  comments: Comments;
  hyperlinks: Hyperlinks;
  annotations: Annotations;
  footnotes: Footnotes;
  headerFooter: HeaderFooter;
}

const DEFAULTS: Options = {
  trackedChanges: 'annotate',
  comments: 'inline',
  hyperlinks: 'inline',
  annotations: 'html',
  footnotes: 'end',
  headerFooter: 'strip',
};

function isPaged(r: MarkdownResult | PagedMarkdownResult): r is PagedMarkdownResult {
  return Array.isArray((r as PagedMarkdownResult).pages);
}

export function App() {
  const [fileName, setFileName] = useState<string>('');
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  const [method, setMethod] = useState<Method>('continuous');
  const [opts, setOpts] = useState<Options>(DEFAULTS);
  const [result, setResult] = useState<MarkdownResult | PagedMarkdownResult | null>(null);
  const [error, setError] = useState<string>('');
  const [dragover, setDragover] = useState(false);

  const handleFile = useCallback(async (file: File) => {
    setError('');
    setResult(null);
    setFileName(file.name);
    const buf = await file.arrayBuffer();
    setBuffer(buf);
  }, []);

  // Editor needs a stable buffer reference; cloning avoids the editor and
  // converter both reading the same ArrayBuffer (the editor's parser detaches it).
  const editorBuffer = useMemo(() => buffer?.slice(0), [buffer]);
  const converterBuffer = useMemo(() => buffer?.slice(0), [buffer]);

  // Re-run conversion whenever buffer or options change.
  useEffect(() => {
    if (!converterBuffer) return;
    let cancelled = false;
    (async () => {
      try {
        setError('');
        const baseOpts = {
          trackedChanges: opts.trackedChanges,
          comments: opts.comments,
          hyperlinks: opts.hyperlinks,
          annotations: opts.annotations,
        };
        if (method === 'paged') {
          const r = await toMarkdownPaged(new Uint8Array(converterBuffer), {
            ...baseOpts,
            footnotes: opts.footnotes,
            headerFooter: opts.headerFooter,
          });
          if (!cancelled) setResult(r);
        } else {
          const r = await toMarkdown(new Uint8Array(converterBuffer), {
            ...baseOpts,
            footnotes: opts.footnotes,
          });
          if (!cancelled) setResult(r);
        }
      } catch (err) {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [converterBuffer, method, opts]);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragover(false);
      const file = e.dataTransfer.files?.[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  return (
    <div className="playground">
      <Toolbar
        fileName={fileName}
        method={method}
        setMethod={setMethod}
        opts={opts}
        setOpts={setOpts}
        result={result}
        onFile={handleFile}
      />
      <div className="panes">
        <div className="pane">
          <div className="pane-header">Rendered DOCX</div>
          <div
            className="pane-body"
            onDragOver={(e) => {
              e.preventDefault();
              setDragover(true);
            }}
            onDragLeave={() => setDragover(false)}
            onDrop={onDrop}
          >
            {editorBuffer ? (
              <DocxEditor
                documentBuffer={editorBuffer}
                showToolbar={false}
                showZoomControl={false}
                showOutlineButton={false}
                readOnly
                key={fileName}
              />
            ) : (
              <div className={`drop-empty ${dragover ? 'dragover' : ''}`}>
                <span className="material-symbols-outlined" style={{ fontSize: 48 }}>
                  upload_file
                </span>
                <div>Drop a .docx file here</div>
                <label>
                  <input
                    type="file"
                    accept=".docx"
                    style={{ display: 'none' }}
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      if (f) handleFile(f);
                    }}
                  />
                  <button
                    type="button"
                    onClick={(e) => (e.currentTarget.previousSibling as HTMLInputElement)?.click()}
                  >
                    choose file
                  </button>
                </label>
              </div>
            )}
          </div>
        </div>
        <div className="pane">
          <div className="pane-header">Markdown output</div>
          <div className="pane-body">
            {error ? (
              <pre className="error">{error}</pre>
            ) : result ? (
              <MarkdownView result={result} />
            ) : (
              <div className="drop-empty">no document loaded</div>
            )}
          </div>
          {result && result.warnings.length > 0 && (
            <div className="warnings">
              <strong>{result.warnings.length} warning(s)</strong>
              <ul>
                {result.warnings.slice(0, 10).map((w: string, i: number) => (
                  <li key={i}>{w}</li>
                ))}
                {result.warnings.length > 10 && <li>… and {result.warnings.length - 10} more</li>}
              </ul>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function MarkdownView({ result }: { result: MarkdownResult | PagedMarkdownResult }) {
  if (isPaged(result)) {
    return (
      <div className="md-pages">
        {result.pages.map((p: { pageNumber: number; markdown: string }) => (
          <div className="md-page" key={p.pageNumber}>
            <div className="md-page-label">Page {p.pageNumber}</div>
            <pre className="md-page-body">{p.markdown}</pre>
          </div>
        ))}
      </div>
    );
  }
  return <pre className="markdown-output">{result.markdown}</pre>;
}

interface ToolbarProps {
  fileName: string;
  method: Method;
  setMethod: (m: Method) => void;
  opts: Options;
  setOpts: (o: Options) => void;
  result: MarkdownResult | PagedMarkdownResult | null;
  onFile: (f: File) => void;
}

function Toolbar({ fileName, method, setMethod, opts, setOpts, result, onFile }: ToolbarProps) {
  const update = <K extends keyof Options>(k: K, v: Options[K]) => setOpts({ ...opts, [k]: v });
  const text = result ? (isPaged(result) ? result.combined : result.markdown) : '';
  const wordCount = text ? text.split(/\s+/).filter(Boolean).length : 0;
  const fmt = (n: number) => n.toLocaleString();
  const stats = result
    ? isPaged(result)
      ? `${result.pages.length} pages · ${fmt(wordCount)} words · ${result.images.size} imgs`
      : `${fmt(wordCount)} words · ${result.images.size} imgs`
    : '';

  return (
    <div className="toolbar">
      <h1>DOCX → MD playground</h1>
      <span className="pill">{fileName || 'no file'}</span>
      <label>
        <input
          type="file"
          accept=".docx"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) onFile(f);
          }}
        />
      </label>

      <label>
        method
        <select value={method} onChange={(e) => setMethod(e.target.value as Method)}>
          <option value="continuous">continuous (toMarkdown)</option>
          <option value="paged">paged (toMarkdownPaged)</option>
        </select>
      </label>

      <label>
        annotations
        <select
          value={opts.annotations}
          onChange={(e) => update('annotations', e.target.value as Annotations)}
        >
          <option value="html">html</option>
          <option value="pandoc">pandoc</option>
          <option value="strip">strip</option>
        </select>
      </label>

      <label>
        trackedChanges
        <select
          value={opts.trackedChanges}
          onChange={(e) => update('trackedChanges', e.target.value as Tracked)}
        >
          <option value="annotate">annotate</option>
          <option value="clean">clean</option>
        </select>
      </label>

      <label>
        comments
        <select
          value={opts.comments}
          onChange={(e) => update('comments', e.target.value as Comments)}
        >
          <option value="inline">inline</option>
          <option value="sidecar">sidecar</option>
          <option value="strip">strip</option>
        </select>
      </label>

      <label>
        hyperlinks
        <select
          value={opts.hyperlinks}
          onChange={(e) => update('hyperlinks', e.target.value as Hyperlinks)}
        >
          <option value="inline">inline</option>
          <option value="reference">reference</option>
        </select>
      </label>

      <label>
        footnotes
        <select
          value={opts.footnotes}
          onChange={(e) => update('footnotes', e.target.value as Footnotes)}
        >
          <option value="end">end</option>
          <option value="inline">inline</option>
        </select>
      </label>

      <label>
        headerFooter
        <select
          value={opts.headerFooter}
          onChange={(e) => update('headerFooter', e.target.value as HeaderFooter)}
          disabled={method !== 'paged'}
        >
          <option value="strip">strip</option>
          <option value="first-page">first-page</option>
          <option value="all">all</option>
        </select>
      </label>

      <span className="spacer" />
      <span className="stats">
        <strong>{stats}</strong>
      </span>
    </div>
  );
}
