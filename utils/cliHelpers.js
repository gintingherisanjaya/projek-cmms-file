/**
 * CLI flags when stdin is non-interactive (--local, samples, Drive limits, etc.)
 */
export function parseArgs(argv = process.argv.slice(2)) {
  const out = {
    local: false,
    samples: null,
    driveSamples: null,
    singleFileId: null,
    driveSingleFileId: null,
    outputToLocal: false,
    boldFunclocDesc: true,
  };

  if (argv.includes('--local')) out.local = true;
  if (argv.includes('--output-local')) out.outputToLocal = true;
  if (argv.includes('--output-drive')) out.outputToLocal = false;

  const si = argv.indexOf('--single-file-id');
  if (si !== -1 && argv[si + 1]) out.singleFileId = argv[si + 1];

  const samp = argv.indexOf('--samples');
  if (samp !== -1 && argv[samp + 1]) {
    const n = parseInt(argv[samp + 1], 10);
    if (!Number.isNaN(n)) out.samples = n;
  }

  const ds = argv.indexOf('--drive-samples');
  if (ds !== -1 && argv[ds + 1]) {
    const n = parseInt(argv[ds + 1], 10);
    if (!Number.isNaN(n)) out.driveSamples = n;
  }

  return out;
}
