"""CANDYTUNE アスキーアートバナー表示"""
from rich.console import Console
from rich.panel import Panel
from rich.text import Text


BANNER_ART = r""" 
   ██████╗ █████╗ ███╗   ██╗██████╗ ██╗   ██╗████████╗██╗   ██╗███╗   ██╗███████╗
  ██╔════╝██╔══██╗████╗  ██║██╔══██╗╚██╗ ██╔╝╚══██╔══╝██║   ██║████╗  ██║██╔════╝
  ██║     ███████║██╔██╗ ██║██║  ██║ ╚████╔╝    ██║   ██║   ██║██╔██╗ ██║█████╗  
  ██║     ██╔══██║██║╚██╗██║██║  ██║  ╚██╔╝     ██║   ██║   ██║██║╚██╗██║██╔══╝  
  ╚██████╗██║  ██║██║ ╚████║██████╔╝   ██║      ██║   ╚██████╔╝██║ ╚████║███████╗
  ╚═════╝╚═╝  ╚═╝╚═╝  ╚═══╝╚═════╝    ╚═╝      ╚═╝    ╚═════╝ ╚═╝  ╚═══╝╚══════╝
"""


def show_banner():
    """起動時にCANDYTUNEバナーを表示"""
    console = Console()
    
    # グラデーション風にカラーリング
    banner_lines = BANNER_ART.strip().split('\n')
    colored_banner = Text()
    
    # ピンク→パープル→ブルーのグラデーション
    colors = [
        "bright_magenta",
        "magenta",
        "bright_blue",
        "blue",
        "bright_cyan",
        "cyan",
    ]
    
    for i, line in enumerate(banner_lines):
        color = colors[i % len(colors)]
        colored_banner.append(line + "\n", style=f"bold {color}")
    
    # パネルで囲んで表示
    panel = Panel(
        colored_banner,
        title="[bold yellow]✨ PDF Converter ✨[/bold yellow]",
        subtitle="[italic cyan]Turn Anything into PDFs. Fast. In Docker.[/italic cyan]",
        border_style="bright_magenta",
        padding=(1, 2),
    )
    
    console.print()
    console.print(panel)
    console.print()

