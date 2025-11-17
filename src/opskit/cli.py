import click

from opskit.sharepoint.cli import sharepoint_cli

@click.group()
@click.option("--debug", is_flag=True, envvar="DEBUG", help="Enable debug mode.")
@click.pass_context
def cli(ctx, debug):
    ctx.ensure_object(dict)
    ctx.obj["DEBUG"] = debug

cli.add_command(sharepoint_cli)