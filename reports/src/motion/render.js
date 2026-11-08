const { chromium } = require('playwright'); const { spawn } = require('child_process');
const M='/tmp/claude-0/-home-user-sisasales/82ad116b-b749-5857-911d-f0676cb5351b/scratchpad/mg', MODE=process.argv[2], FPS=30;
(async()=>{
  const b=await chromium.launch(); const pg=await b.newPage({viewport:{width:1920,height:1080},deviceScaleFactor:1});
  await pg.goto('file://'+M+'/build/index.html'); await pg.evaluate(()=>document.fonts.ready); await pg.waitForTimeout(800);
  const total=await pg.evaluate('TOTAL');
  if(MODE==='test'){ for(const t of process.argv[3].split(',').map(Number)){ await pg.evaluate('render('+t+')'); await pg.screenshot({path:M+'/'+process.argv[4]+'/t'+t.toFixed(2).padStart(6,'0')+'.png'}); } }
  else { const ff=spawn('ffmpeg',['-y','-loglevel','error','-f','image2pipe','-framerate',String(FPS),'-c:v','mjpeg','-i','-','-c:v','libx264','-preset','veryfast','-crf','18','-pix_fmt','yuv420p',M+'/silent.mp4'],{stdio:['pipe','inherit','inherit']});
    const n=Math.floor(total*FPS);
    for(let i=0;i<n;i++){ await pg.evaluate('render('+(i/FPS)+')'); const buf=await pg.screenshot({type:'jpeg',quality:93}); if(!ff.stdin.write(buf)) await new Promise(r=>ff.stdin.once('drain',r)); }
    ff.stdin.end(); await new Promise(r=>ff.on('close',r)); }
  await b.close();
})();
