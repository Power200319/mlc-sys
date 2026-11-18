from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_GET, require_POST

from .grade_config import get_available_grades, get_grade_label
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