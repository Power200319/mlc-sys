from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_GET, require_POST

from .grade_config import get_available_grades, get_grade_label
from .early_years import consume_early_years_output, early_years_profiles, process_early_years as process_early_years_workbook
from .services import ProcessingError, consume_output, process_gradebooks


@require_GET
def health(request):
    return JsonResponse({"status": "ok"})


@require_GET
def grades(request):
    """List available grades with labels."""
    return JsonResponse({
        "grades": [
            {"value": g, "label": get_grade_label(g)}
            for g in get_available_grades()
        ],
    })


@require_GET
def early_years_profiles_view(request):
    """Expose the separate K1--K3 workflow without mixing it with E-Gradebook."""
    return JsonResponse({"profiles": early_years_profiles()})


@csrf_exempt
@require_POST
def process_early_years(request):
    upload = request.FILES.get("e_progress")
    try:
        return JsonResponse(
            process_early_years_workbook(upload, request.POST.get("grade", ""), request.POST.get("term", "")),
            status=201,
        )
    except ProcessingError as error:
        return JsonResponse({"error": str(error)}, status=400)
    except Exception:
        return JsonResponse({"error": "Early Years processing failed unexpectedly. No output has been marked successful."}, status=500)


@csrf_exempt
@require_POST
def process(request):
    uploads = request.FILES.getlist("gradebooks") or request.FILES.getlist("gradebook")
    if not uploads:
        return JsonResponse({"error": "Choose one or more .xlsx E-gradebook files."}, status=400)
    try:
        grade = int(request.POST.get("grade", "6"))
    except (TypeError, ValueError):
        grade = 6
    try:
        return JsonResponse(
            process_gradebooks(
                uploads,
                request.POST.get("period", ""),
                request.FILES.get("scoring_workbook"),
                request.POST.get("print_stage", ""),
                grade,
            ),
            status=201,
        )
    except ProcessingError as error:
        return JsonResponse({"error": str(error)}, status=400)
    except Exception:
        return JsonResponse({"error": "Processing failed unexpectedly. No output has been marked successful."}, status=500)


@require_GET
def download(request, job_id, kind):
    try:
        grade = int(request.GET.get("grade", "6"))
    except (TypeError, ValueError):
        grade = 6
    try:
        content, filename = consume_output(job_id, kind, grade)
    except ProcessingError as error:
        return JsonResponse({"error": str(error)}, status=404)
    response = HttpResponse(content)
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    response["Content-Length"] = str(len(content))
    response["Cache-Control"] = "no-store"
    return response


@require_GET
def download_early_years(request, job_id):
    try:
        content, filename = consume_early_years_output(job_id)
    except ProcessingError as error:
        return JsonResponse({"error": str(error)}, status=404)
    response = HttpResponse(content, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    response["Content-Length"] = str(len(content))
    response["Cache-Control"] = "no-store"
    return response
