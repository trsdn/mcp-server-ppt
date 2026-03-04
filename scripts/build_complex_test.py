import asyncio, json, os
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

DOTNET = r'C:\Users\torstenmahr\GitHub\mcp-server-ppt\src\PptMcp.McpServer\bin\Debug\net9.0-windows\PptMcp.McpServer.dll'
FILE = r'C:\Users\torstenmahr\OneDrive - Microsoft\Desktop\PptMcp-ComplexTest.pptx'

if os.path.exists(FILE):
    os.remove(FILE)

async def ct(session, name, args):
    r = await session.call_tool(name, args)
    text = r.content[0].text if r.content else '{}'
    parsed = json.loads(text)
    action = args.get('action', '?')
    ok = parsed.get('success', False)
    status = "OK" if ok else "FAIL: " + str(parsed.get('errorMessage', ''))[:100]
    print(f'  {name}.{action} -> {status}', flush=True)
    return parsed

async def main():
    server = StdioServerParameters(command='dotnet', args=[DOTNET])
    async with stdio_client(server) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            print('=== Creating complex test presentation ===', flush=True)

            # 1. Create empty file
            r = await ct(session, 'file', {'action': 'create', 'path': FILE, 'show': True})
            sid = r.get('session_id', '')
            print(f'  Session: {sid}', flush=True)

            # === SLIDE 1: Title Slide ===
            await ct(session, 'slide', {'action': 'create', 'session_id': sid, 'position': 0, 'layout_name': 'Title Slide'})
            # Find placeholder names
            s1 = await ct(session, 'shape', {'action': 'list', 'session_id': sid, 'slide_index': 1})
            s1_names = [s['name'] for s in s1.get('shapes', [])]
            print(f'  Slide 1 shapes: {s1_names}', flush=True)
            if len(s1_names) >= 2:
                await ct(session, 'text', {'action': 'set', 'session_id': sid, 'slide_index': 1, 'shape_name': s1_names[0], 'text': 'PptMcp Complex Demo'})
                await ct(session, 'text', {'action': 'set', 'session_id': sid, 'slide_index': 1, 'shape_name': s1_names[1], 'text': 'Testing All COM Components\nTable \u2022 Chart \u2022 Animation \u2022 Design \u2022 Shapes'})
                await ct(session, 'text', {'action': 'format', 'session_id': sid, 'slide_index': 1, 'shape_name': s1_names[0], 'font_name': 'Segoe UI', 'font_size': 44, 'bold': True, 'color': '#1B3A5C'})

            # === SLIDE 2: Data Table ===
            await ct(session, 'slide', {'action': 'create', 'session_id': sid, 'position': 0, 'layout_name': 'Blank'})
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 2, 'left': 50, 'top': 20, 'width': 620, 'height': 60, 'text': 'Quarterly Financial Summary'})
            await ct(session, 'slidetable', {'action': 'create', 'session_id': sid, 'slide_index': 2, 'rows': 5, 'columns': 4, 'left': 50, 'top': 100, 'width': 620, 'height': 250})

            # Find table shape name dynamically
            s2 = await ct(session, 'shape', {'action': 'list', 'session_id': sid, 'slide_index': 2})
            tbl = next((s['name'] for s in s2.get('shapes', []) if s.get('hasTable')), '')
            tb2 = [s['name'] for s in s2.get('shapes', []) if s.get('hasTextFrame') and not s.get('hasTable')]
            print(f'  Table: {tbl}, Title: {tb2}', flush=True)

            if tb2:
                await ct(session, 'text', {'action': 'format', 'session_id': sid, 'slide_index': 2, 'shape_name': tb2[0], 'font_name': 'Segoe UI', 'font_size': 28, 'bold': True, 'color': '#2E4057'})

            if tbl:
                # Headers
                for ci, h in enumerate(['Quarter', 'Revenue', 'Costs', 'Profit'], 1):
                    await ct(session, 'slidetable', {'action': 'write-cell', 'session_id': sid, 'slide_index': 2, 'shape_name': tbl, 'row': 1, 'column': ci, 'value': h})
                # Data
                for ri, row in enumerate([['Q1','$1.2M','$800K','$400K'],['Q2','$1.5M','$900K','$600K'],['Q3','$1.8M','$950K','$850K'],['Q4','$2.1M','$1.0M','$1.1M']], 2):
                    for ci, val in enumerate(row, 1):
                        await ct(session, 'slidetable', {'action': 'write-cell', 'session_id': sid, 'slide_index': 2, 'shape_name': tbl, 'row': ri, 'column': ci, 'value': val})

            # === SLIDE 3: Chart ===
            await ct(session, 'slide', {'action': 'create', 'session_id': sid, 'position': 0, 'layout_name': 'Blank'})
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 3, 'left': 80, 'top': 20, 'width': 560, 'height': 60, 'text': 'Revenue Trend Analysis'})
            await ct(session, 'chart', {'action': 'create', 'session_id': sid, 'slide_index': 3, 'chart_type': 51, 'left': 80, 'top': 100, 'width': 560, 'height': 350})

            s3 = await ct(session, 'shape', {'action': 'list', 'session_id': sid, 'slide_index': 3})
            chart_name = next((s['name'] for s in s3.get('shapes', []) if s.get('hasChart')), '')
            print(f'  Chart: {chart_name}', flush=True)
            if chart_name:
                await ct(session, 'chart', {'action': 'set-title', 'session_id': sid, 'slide_index': 3, 'shape_name': chart_name, 'title': 'Revenue by Quarter'})

            # === SLIDE 4: Shapes showcase ===
            await ct(session, 'slide', {'action': 'create', 'session_id': sid, 'position': 0, 'layout_name': 'Blank'})
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 4, 'left': 50, 'top': 20, 'width': 620, 'height': 50, 'text': 'Shape Showcase'})
            for stype, l, t, w, h in [(1,50,120,150,100),(9,240,120,150,100),(5,430,120,150,100),(4,50,280,150,120),(56,240,280,150,120)]:
                await ct(session, 'shape', {'action': 'add-shape', 'session_id': sid, 'slide_index': 4, 'auto_shape_type': stype, 'left': l, 'top': t, 'width': w, 'height': h})
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 4, 'left': 50, 'top': 430, 'width': 620, 'height': 50, 'text': '5 AutoShape types rendered via COM Interop'})

            # Format title
            s4 = await ct(session, 'shape', {'action': 'list', 'session_id': sid, 'slide_index': 4})
            first_tb4 = next((s['name'] for s in s4.get('shapes', []) if s.get('hasTextFrame') and 'Shape Showcase' in (s.get('text',''))), None)
            if first_tb4:
                await ct(session, 'text', {'action': 'format', 'session_id': sid, 'slide_index': 4, 'shape_name': first_tb4, 'font_name': 'Segoe UI', 'font_size': 28, 'bold': True})

            # === SLIDE 5: Animation demo ===
            await ct(session, 'slide', {'action': 'create', 'session_id': sid, 'position': 0, 'layout_name': 'Blank'})
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 5, 'left': 50, 'top': 20, 'width': 620, 'height': 50, 'text': 'Animation Demo'})
            await ct(session, 'shape', {'action': 'add-shape', 'session_id': sid, 'slide_index': 5, 'auto_shape_type': 1, 'left': 100, 'top': 150, 'width': 200, 'height': 80})
            await ct(session, 'shape', {'action': 'add-shape', 'session_id': sid, 'slide_index': 5, 'auto_shape_type': 9, 'left': 400, 'top': 150, 'width': 200, 'height': 80})
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 5, 'left': 100, 'top': 300, 'width': 500, 'height': 80, 'text': 'These shapes have entrance animations!\nClick to trigger Fade, Fly, and Appear effects.'})

            s5 = await ct(session, 'shape', {'action': 'list', 'session_id': sid, 'slide_index': 5})
            all5 = [s['name'] for s in s5.get('shapes', [])]
            print(f'  Slide 5 shapes: {all5}', flush=True)
            # Animate: skip title textbox (index 0), animate shapes 1,2,3
            if len(all5) >= 4:
                await ct(session, 'animation', {'action': 'add', 'session_id': sid, 'slide_index': 5, 'shape_name': all5[1], 'effect_type': 10, 'trigger_type': 1})  # Fade OnClick
                await ct(session, 'animation', {'action': 'add', 'session_id': sid, 'slide_index': 5, 'shape_name': all5[2], 'effect_type': 2, 'trigger_type': 2})   # Fly WithPrev
                await ct(session, 'animation', {'action': 'add', 'session_id': sid, 'slide_index': 5, 'shape_name': all5[3], 'effect_type': 1, 'trigger_type': 3})   # Appear AfterPrev

            anims = await ct(session, 'animation', {'action': 'list', 'session_id': sid, 'slide_index': 5})
            print(f'  Animations: {len(anims.get("animations", []))}', flush=True)

            # === SLIDE 6: Design & Theme info ===
            await ct(session, 'slide', {'action': 'create', 'session_id': sid, 'position': 0, 'layout_name': 'Blank'})
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 6, 'left': 50, 'top': 20, 'width': 620, 'height': 50, 'text': 'Design & Theme Info'})

            colors = await ct(session, 'design', {'action': 'get-colors', 'session_id': sid, 'design_index': 1})
            ctxt = 'Theme Color Palette:\n'
            for n, v in list(colors.get('colors', {}).items())[:8]:
                ctxt += f'  {n}: {v}\n'
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 6, 'left': 50, 'top': 100, 'width': 350, 'height': 300, 'text': ctxt})

            designs = await ct(session, 'design', {'action': 'list', 'session_id': sid})
            dtxt = 'Designs:\n'
            for d in designs.get('designs', []):
                dtxt += f'  {d["name"]} ({d["layoutCount"]} layouts)\n'
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 6, 'left': 420, 'top': 100, 'width': 280, 'height': 300, 'text': dtxt})

            # === SLIDE 7: Transitions summary ===
            await ct(session, 'slide', {'action': 'create', 'session_id': sid, 'position': 0, 'layout_name': 'Blank'})
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 7, 'left': 50, 'top': 20, 'width': 620, 'height': 50, 'text': 'Transition Effects Applied'})
            for si, eff, dur in [(2,2145,1.5),(3,2026,1.0),(4,2037,0.8),(5,2050,1.2)]:
                await ct(session, 'transition', {'action': 'set', 'session_id': sid, 'slide_index': si, 'effect_type': eff, 'duration': dur, 'advance_on_click': True})
            await ct(session, 'shape', {'action': 'add-textbox', 'session_id': sid, 'slide_index': 7, 'left': 50, 'top': 120, 'width': 620, 'height': 250,
                'text': 'Slide 2: Push transition (1.5s)\nSlide 3: Wipe transition (1.0s)\nSlide 4: Split transition (0.8s)\nSlide 5: Fade transition (1.2s)\n\nNavigate slides to see effects!'})

            # === NOTES ===
            await ct(session, 'notes', {'action': 'set', 'session_id': sid, 'slide_index': 1, 'text': 'Title slide of the PptMcp complex demo.'})
            await ct(session, 'notes', {'action': 'set', 'session_id': sid, 'slide_index': 2, 'text': 'Financial data table. Data entered cell-by-cell via COM.'})
            await ct(session, 'notes', {'action': 'set', 'session_id': sid, 'slide_index': 5, 'text': 'Animation demo - Fade, Fly, Appear effects via COM.'})

            # === WINDOW INFO ===
            win = await ct(session, 'window', {'action': 'get-info', 'session_id': sid})
            print(f'  Window: {win.get("windowStateName", "?")}', flush=True)

            # === FINAL SUMMARY ===
            slides = await ct(session, 'slide', {'action': 'list', 'session_id': sid})
            print(f'  Total slides: {len(slides.get("slides", []))}', flush=True)
            for s in slides.get('slides', []):
                print(f'    Slide {s["slideIndex"]}: {s["layoutName"]} ({s["shapeCount"]} shapes, notes={s.get("hasNotes",False)}, anim={s.get("hasAnimations",False)})', flush=True)

            # === SAVE & CLOSE ===
            await ct(session, 'file', {'action': 'save', 'session_id': sid})
            await ct(session, 'file', {'action': 'close', 'session_id': sid, 'save': False})

            print('\n=== DONE! Complex presentation saved to Desktop ===', flush=True)

asyncio.run(main())
