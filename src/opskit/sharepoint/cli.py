import click

@click.group()
@click.pass_context
def sharepoint_cli(ctx):
    """SharePoint related commands."""
    pass

@click.command()
def upload_cmd(ctx):
    """Upload files to SharePoint."""
    click.echo("Uploading files to SharePoint...")