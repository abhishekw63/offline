"""
Django views for the Offline / GT Mass Dump Generator.

Routes:
    /offline/                → OfflineDashboardView (department landing)
    /offline/gt-mass-dump/   → IndexView (upload form)
    /offline/process/        → ProcessFilesView (file processing + download)
"""

from django.views.generic import TemplateView, View
from django.contrib.auth.mixins import LoginRequiredMixin
from django.http import HttpResponse, JsonResponse
from datetime import datetime

from .utils import GTMassAutomation


class OfflineDashboardView(LoginRequiredMixin, TemplateView):
    """Department landing page — links to GT Mass Dump and future tools."""
    template_name = 'offline/dashboard.html'


class IndexView(LoginRequiredMixin, TemplateView):
    """Upload form for GT Mass Dump Generator."""
    template_name = 'offline/index.html'


class ProcessFilesView(LoginRequiredMixin, View):
    """
    Process uploaded Excel files and return the output dump.

    POST /offline/process/
        Body: multipart form with 'files' field (multiple .xlsx/.xls)

    Success response:
        Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
        Content-Disposition: attachment; filename="gt_mass_dump_DDMMYYYY.xlsx"
        X-GT-Summary: JSON string with processing summary

    Error response:
        JSON { "error": "...", "details": {...} }
    """

    def post(self, request, *args, **kwargs):
        files = request.FILES.getlist('files')

        if not files:
            return JsonResponse(
                {"error": "No files selected"},
                status=400,
            )

        # ── Process all files ──
        automation = GTMassAutomation()
        result = automation.process_files(files)

        # ── Export to memory ──
        output_buffer = automation.exporter.export_to_memory(result)

        if output_buffer is None:
            return JsonResponse(
                {
                    "error": "No valid data found in selected files",
                    "details": {
                        "attempted": len(result.attempted_files),
                        "failed": len(result.failed_files),
                        "failures": [
                            {"file": f, "reason": r}
                            for f, r in result.failed_files
                        ],
                    },
                },
                status=400,
            )

        # ── Build response ──
        today = datetime.now().strftime("%d%m%Y")
        filename = f"gt_mass_dump_{today}.xlsx"

        # Summary stats for the frontend to display
        unique_sos = len({r.so_number for r in result.rows})
        missing_loc = len({r.so_number for r in result.rows if not r.location_code})

        response = HttpResponse(
            output_buffer.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'

        # Pass summary as custom headers so JS can display stats
        response['X-GT-Attempted'] = str(len(result.attempted_files))
        response['X-GT-Rows'] = str(len(result.rows))
        response['X-GT-SOs'] = str(unique_sos)
        response['X-GT-Failed'] = str(len(result.failed_files))
        response['X-GT-Warnings'] = str(len(result.warned_files))
        response['X-GT-MissingLocation'] = str(missing_loc)

        # Expose custom headers to JavaScript
        response['Access-Control-Expose-Headers'] = (
            'X-GT-Attempted, X-GT-Rows, X-GT-SOs, '
            'X-GT-Failed, X-GT-Warnings, X-GT-MissingLocation'
        )

        return response