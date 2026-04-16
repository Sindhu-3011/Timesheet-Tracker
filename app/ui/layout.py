"""app/ui/layout.py — Page chrome: nav bar, sidebar, render_page."""
import html
from .styles import PAGE_STYLE
from ..database.user_crud import get_feature_flags_by_email, get_feature_flags_by_username
from ..config import normalize_screen_access

def NAV(display_name="Visitor"):
    brand = (
        '<div class="brand"><div class="logo">VS</div>'
        '<div><div style="font-size:14px">Vaisesika</div>'
        '<div style="font-size:12px;color:#9fb3d8;margin-top:2px">Timesheet Portal</div></div></div>'
    )
    if display_name == "Visitor":
        return (
            '<header class="app-header">'
            f'{brand}'
            '<div class="header-right"><a class="btn outline small" style="padding:8px 12px;font-size:13px;margin-right:8px" href="/help"><i class="fa fa-circle-question"></i> Help</a>Welcome</div>'
            '</header>'
        )
    return (
        '<header class="app-header">'
        f'{brand}'
        f'<div class="header-right">Signed in as <strong>{html.escape(display_name)}</strong>'
        f'<a class="btn outline" style="padding:8px 12px;font-size:13px;margin-left:12px" href="/help"><i class="fa fa-circle-question"></i> Help</a><a class="btn secondary" style="padding:8px 12px;font-size:13px;margin-left:12px" href="/logout">Logout</a>'
        '</div></header>'
    )






def build_sidebar_for_role(role, screen_access="BOTH", can_email=0, can_export=0, can_help=1):
    # Admin always has full access
    if role == "Admin":
        screen_access = "BOTH"
        can_email, can_export, can_help = 1, 1, 1
    # Admin always has full access
    if role == "Admin":
        screen_access = "BOTH"
        can_email, can_export, can_help = 1, 1, 1

    screen_access = normalize_screen_access(screen_access)

    module_links = []
    if role == "Admin" or screen_access == "BOTH":
        module_links = [
            '<a href="/upload/PPM" title="PPM screen prints">PPM</a>',
            '<a href="/upload/NTT" title="NTT screen prints">NTT</a>',
            '<a href="/upload/EMAIL" title="Email screen prints">EMAIL</a>',
        ]
    elif screen_access == "PPM":
        module_links = ['<a href="/upload/PPM" title="PPM screen prints">PPM</a>']
    elif screen_access == "NTT":
        module_links = ['<a href="/upload/NTT" title="NTT screen prints">NTT</a>']
    elif screen_access == "EMAIL":
        module_links = ['<a href="/upload/EMAIL" title="Email screen prints">EMAIL</a>']

    screen_module_dropdown = (
        '<div class="nav-item"><i class="fas fa-image"></i> Screen Print Module'
        '<div class="dropdown-content">'
        + "".join(module_links) +
        '</div></div>'
    )

    parts = ['<aside class="sidebar">']

    if role == "Admin":
        parts.append(
            '<div class="nav-item"><i class="fas fa-users"></i> User Management'
            '<div class="dropdown-content">'
            '<a href="/user-management/create">Create User</a>'
            '<a href="/user-management/list">User List</a><a href="/user-management/create-fields">Create Fields</a>'
            '</div></div>'
        )

    parts.append(screen_module_dropdown)

    if can_email:
        parts.append('<a class="nav-item" href="/email-settings"><i class="fas fa-envelope"></i> Email Settings</a>')
    if can_export:
        parts.append('<a class="nav-item" href="/export"><i class="fas fa-file-export"></i> Export Data</a>')
    parts.append('</aside>')
    return "".join(parts)


def render_page(content, display_name="Visitor", role=None, screen_access="BOTH"):
    if display_name == "Visitor":
        return (
            f"<html><head>{PAGE_STYLE}</head>"
            f"<body>{NAV(display_name)}"
            f"<main class='content' style='margin-top:64px;padding:28px'>{content}</main>"
            f"</body></html>"
        )

    if role == "Admin":
        can_email, can_export, can_help = (1, 1, 1)
    else:
        if isinstance(display_name, str) and "@" in display_name:
            can_email, can_export, can_help = get_feature_flags_by_email(display_name)
        else:
            can_email, can_export, can_help = get_feature_flags_by_username(display_name)

    sidebar_html = build_sidebar_for_role(role or "", screen_access or "BOTH",
                                         can_email=can_email, can_export=can_export, can_help=can_help)
    return (
        f"<html><head>{PAGE_STYLE}</head>"
        f"<body>{NAV(display_name)}"
        f"<div class='layout'>{sidebar_html}<main class='content'>{content}</main></div>"
        f"</body></html>"
    )

