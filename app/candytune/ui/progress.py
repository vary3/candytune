"""カラフルなプログレス表示"""
from pathlib import Path
from typing import Optional

from rich.console import Console
from rich.progress import (
    Progress,
    SpinnerColumn,
    TextColumn,
    BarColumn,
    TaskProgressColumn,
    TimeRemainingColumn,
    TimeElapsedColumn,
)
from rich.table import Table
from rich.panel import Panel
from rich import box


class ConversionProgress:
    """変換プロセスのプログレス表示"""
    
    def __init__(self):
        self.console = Console()
        self.converted = 0
        self.errors = []
        
    def create_progress_bar(self, total: int) -> Progress:
        """プログレスバーを作成"""
        return Progress(
            SpinnerColumn(spinner_name="dots12", style="cyan"),
            TextColumn("[bold blue]{task.description}"),
            BarColumn(
                complete_style="green",
                finished_style="bright_green",
                pulse_style="yellow",
            ),
            TaskProgressColumn(),
            TextColumn("•"),
            TimeElapsedColumn(),
            TextColumn("•"),
            TimeRemainingColumn(),
            console=self.console,
        )
    
    def print_converting(self, file_path: Path, output_path: Path):
        """変換中のファイル情報を表示"""
        self.console.print(
            f"  [green]✓[/green] [cyan]{file_path.name}[/cyan] → "
            f"[yellow]{output_path.name}[/yellow]"
        )
    
    def print_error(self, file_path: Path, error_msg: str):
        """エラー情報を表示"""
        self.console.print(
            f"  [red]✗[/red] [red]{file_path.name}[/red]: {error_msg}",
            style="dim",
        )
    
    def print_warning(self, file_path: Path, warning_msg: str):
        """Display warning information (e.g., fallback)"""
        self.console.print(
            f"  [yellow]⚠[/yellow] [dim]{file_path.name}: {warning_msg}[/dim]",
        )
    
    def show_summary(self, total: int, converted: int, errors: list):
        """Display conversion summary"""
        self.console.print()
        
        # Summary table
        table = Table(
            title="[bold cyan]Conversion Summary[/bold cyan]",
            box=box.ROUNDED,
            show_header=True,
            header_style="bold magenta",
        )
        
        table.add_column("Item", style="cyan", justify="left")
        table.add_column("Count", style="yellow", justify="right")
        table.add_column("Status", justify="center")
        
        table.add_row(
            "Total Files",
            str(total),
            "📁",
        )
        table.add_row(
            "Succeeded",
            str(converted),
            "[green]✓[/green]" if converted > 0 else "-",
        )
        table.add_row(
            "Failed",
            str(len(errors)),
            "[red]✗[/red]" if len(errors) > 0 else "[green]✓[/green]",
        )
        
        self.console.print(table)
        
        # Error details
        if errors:
            self.console.print()
            error_panel = Panel(
                self._format_errors(errors),
                title="[bold red]Error Details[/bold red]",
                border_style="red",
                padding=(1, 2),
            )
            self.console.print(error_panel)
        else:
            self.console.print()
            self.console.print(
                "[bold green]🎉 All files converted successfully![/bold green]",
                justify="center",
            )
        
        self.console.print()
    
    def _format_errors(self, errors: list, max_display: int = 50) -> str:
        """Format error list"""
        lines = []
        for i, (path, msg) in enumerate(errors[:max_display], 1):
            lines.append(f"{i}. {path.name}: {msg}")
        
        if len(errors) > max_display:
            lines.append(f"\n... and {len(errors) - max_display} more errors")
        
        lines.append("\n")
        lines.append("💡 Troubleshooting:")
        lines.append("  • Check if the file is corrupted")
        lines.append("  • Try opening and re-saving the file manually")
        lines.append("  • Rename files with special characters")
        lines.append("  • Increase mem_limit in docker-compose.yml if out of memory")
        
        return "\n".join(lines)
    
    def print_config(self, input_dir: Path, output_dir: Path, image_dpi: int, flatten: bool):
        """Display configuration"""
        config_table = Table(
            box=box.SIMPLE,
            show_header=False,
            padding=(0, 2),
        )
        
        config_table.add_column("Key", style="cyan bold")
        config_table.add_column("Value", style="yellow")
        
        config_table.add_row("Input Directory", str(input_dir))
        config_table.add_row("Output Directory", str(output_dir))
        config_table.add_row("Image DPI", str(image_dpi))
        config_table.add_row("Flatten", "Enabled" if flatten else "Disabled")
        
        panel = Panel(
            config_table,
            title="[bold blue]Configuration[/bold blue]",
            border_style="blue",
            padding=(1, 2),
        )
        
        self.console.print(panel)
        self.console.print()

