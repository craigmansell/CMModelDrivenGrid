const fs=require('node:fs');
const path=require('node:path');
const esbuild=require(require.resolve('esbuild',{paths:[path.dirname(require.resolve('tsx/package.json'))]}));
const root=path.resolve(__dirname,'..'),out=path.join(root,'out','spacing-demo');
async function start(){
    fs.mkdirSync(out,{recursive:true});
    fs.writeFileSync(path.join(out,'index.html'),`<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Quotation grid spacing harness</title><link rel="stylesheet" href="/bundle.css"><style>body{margin:0;font:14px 'Segoe UI',sans-serif;color:#242424;background:#f5f5f5}header,.tools{padding:12px 18px;display:flex;align-items:center;gap:16px;flex-wrap:wrap}header{background:white;justify-content:space-between}button{font:inherit;padding:5px 10px}main{margin:0 18px 20px;background:white;padding:10px;border:1px solid #ddd;border-radius:6px}.commands{display:flex;align-items:center;justify-content:space-between;padding-bottom:8px}.next-section{margin-top:12px;padding:12px;background:#f0f0f0}</style></head><body><div id="spacing-root"></div><script src="/bundle.js"></script></body></html>`);
    const context=await esbuild.context({absWorkingDir:root,entryPoints:['scripts/quotation-spacing-demo.tsx'],bundle:true,outfile:path.join(out,'bundle.js'),platform:'browser',target:'es2020',define:{'process.env.NODE_ENV':'"development"'},logLevel:'info'});
    await context.watch();const server=await context.serve({host:'127.0.0.1',port:8190,servedir:out});
    console.log(`Spacing harness: http://127.0.0.1:${server.port}/ (sample data only)`);
    process.on('SIGINT',()=>{void context.dispose();});process.on('SIGTERM',()=>{void context.dispose();});
}
start().catch(error=>{console.error(error);process.exitCode=1;});
