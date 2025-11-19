import click

import os

from opskit.sharepoint.client import (
    SharePointClient,
    iter_local,
    count_files,
)

@click.group(context_settings=dict(help_option_names=["-h", "--help"]))
@click.option(
    "--tenant-id",
    envvar="SP_TENANT_ID",
    required=True,
    help="Azure AD Tenant ID (GUID). Env: SP_TENANT_ID",
)
@click.option(
    "--client-id",
    envvar="SP_CLIENT_ID",
    required=True,
    help="App registration Client ID. Env: SP_CLIENT_ID",
)
@click.option(
    "--client-secret",
    envvar="SP_CLIENT_SECRET",
    required=True,
    help="App registration Client Secret. Env: SP_CLIENT_SECRET",
)
@click.option(
    "--site-url",
    envvar="SP_SITE_URL",
    required=True,
    help="SharePoint site URL, e.g. https://contoso.sharepoint.com/sites/Engineering. Env: SP_SITE_URL",
)
@click.option(
    "--library",
    envvar="SP_LIBRARY",
    default="Documents",
    show_default=True,
    help="Document library display name. Env: SP_LIBRARY",
)
@click.option("-v", "--verbose", is_flag=True, help="Verbose output.")
@click.pass_context
def sharepoint_cli(ctx, tenant_id, client_id, client_secret, site_url, library, verbose):
    """SharePoint related commands."""
    pass

@cli.command("upload")
@click.option(
    "--local-folder",
    type=click.Path(exists=True, file_okay=False),
    required=True,
    help="Local folder to upload.",
)
@click.option(
    "--target-folder",
    envvar="SP_TARGET_FOLDER",
    default="",
    show_default=True,
    help="Path inside the document library (created if missing). Env: SP_TARGET_FOLDER",
)
@click.option(
    "--conflict-behavior",
    type=click.Choice(["replace", "rename", "fail"]),
    default="replace",
    show_default=True,
    help="Behavior when file already exists.",
)
@click.option(
    "--small-upload-max",
    type=int,
    default=4 * 1024 * 1024,
    show_default=True,
    help="Threshold (bytes) for simple upload.",
)
@click.option(
    "--chunk-size",
    type=int,
    default=8 * 1024 * 1024,
    show_default=True,
    help="Chunk size (bytes) for large files.",
)
@click.option(
    "--retry-max",
    type=int,
    default=5,
    show_default=True,
    help="Max retries for retryable HTTP errors.",
)
@click.option(
    "--retry-backoff",
    type=float,
    default=2.0,
    show_default=True,
    help="Exponential backoff base for retries.",
)
@click.option(
    "--dry-run",
    is_flag=True,
    help="Validate and show what would be uploaded, without sending anything.",
)
@click.pass_context
def upload_cmd(ctx, local_folder, target_folder, conflict_behavior, small_upload_max, chunk_size, retry_max, retry_backoff, dry_run):
    """Upload files to SharePoint."""
    """
    Upload a LOCAL FOLDER recursively to a SharePoint document library.
    """
    sp: SharePointClient = ctx.obj["sp_client"]
    verbose: bool = ctx.obj["verbose"]

    local_folder = os.path.abspath(local_folder)
    if not os.path.isdir(local_folder):
        raise click.ClickException(f"Local folder does not exist: {local_folder}")

    if verbose:
        click.echo(
            f"[upload] target=/{target_folder} folder={local_folder}",
            err=True,
        )
        click.echo(
            f"[upload] conflict={conflict_behavior}, small_max={small_upload_max}, "
            f"chunk={chunk_size}, retries={retry_max}, backoff={retry_backoff}",
            err=True,
        )
        if dry_run:
            click.echo("[upload] DRY RUN — no changes will be made", err=True)

    # Ensure target folder exists and get its path
    target_item = sp.ensure_folder_path(target_folder, conflict_behavior)
    target_path = sp.get_server_relative_path(target_item)

    click.echo(
        f"Uploading from '{local_folder}' -> {sp.site_url} | {sp.library}:/"
        f"{'/' + target_path if target_path else ''}"
    )
    total_files = count_files(local_folder)
    if total_files == 0:
        click.echo("Nothing to upload (folder is empty).")
        return

    with click.progressbar(length=total_files, label="Uploading files") as bar:
        for dirpath, _, filenames in iter_local(local_folder):
            rel_dir = os.path.relpath(dirpath, local_folder)
            rel_dir = "" if rel_dir == "." else rel_dir.replace("\\", "/")

            folder_path_in_drive = target_path
            if rel_dir:
                folder_path_in_drive = f"{target_path}/{rel_dir}".strip("/")

            # Ensure each subfolder exists
            sp.ensure_folder_path(folder_path_in_drive, conflict_behavior)

            for fname in filenames:
                local_fp = os.path.join(dirpath, fname)
                rel_fp = os.path.relpath(local_fp, local_folder)
                if verbose or dry_run:
                    click.echo(f"→ {rel_fp}  ->  /{folder_path_in_drive}/{fname}")
                sp.upload_file(
                    local_file_path=local_fp,
                    dest_folder_path=folder_path_in_drive,
                    small_upload_max=small_upload_max,
                    chunk_size=chunk_size,
                    conflict_behavior=conflict_behavior,
                    retry_max=retry_max,
                    retry_backoff=retry_backoff,
                    dry_run=dry_run,
                )
                bar.update(1)

    click.echo("Done.")