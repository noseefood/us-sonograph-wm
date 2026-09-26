"""Convert a .pptx (with its animations) into web players under static/anim/.

    python tools/pptx2web/convert.py refs/8190.pptx

Requires Windows + PowerPoint (used to render shapes to SVG), Python with lxml / Pillow / numpy,
and ffmpeg on PATH for slide videos. Steps:
  1. unzip the deck into tools/pptx2web/_work/u
  2. dump.ps1    -> _work/shapes.json   (shape ids, z-order, geometry from PowerPoint)
  3. plan.py     -> _work/plan.json     (timeline per shape, grouping into animation layers)
  4. export.ps1  -> _work/exp/*.svg     (one SVG per layer, rendered by PowerPoint)
  5. assemble.py -> static/anim/slides/slideNN.js + static/anim/media/*
Existing videos in static/anim/media are reused; delete them to re-encode.
"""
import os, shutil, subprocess, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.environ.setdefault('PPTX2WEB_WORK', os.path.join(HERE, '_work'))


def ps(script, pptx):
    subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
                    os.path.join(HERE, script), '-Pptx', pptx, '-Work', WORK], check=True)


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    pptx = os.path.abspath(sys.argv[1])
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK)
    zipfile.ZipFile(pptx).extractall(os.path.join(WORK, 'u'))
    ps('dump.ps1', pptx)
    subprocess.run([sys.executable, os.path.join(HERE, 'plan.py')], check=True)
    ps('export.ps1', pptx)
    subprocess.run([sys.executable, os.path.join(HERE, 'assemble.py')], check=True)


if __name__ == '__main__':
    main()
