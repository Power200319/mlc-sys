from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_GET, require_POST

from .services import ProcessingError, consume_output, process_gradebooks


@require_GET
def health(request):
    return JsonResponse({"status": "ok"})


@csrf_exempt
@require_POST
def process(request):
    uploads = request.FILES.getlist("gradebooks") or request.FILES.getlist("gradebook")
    if not uploads:
        return JsonResponse({"error": "Choose one or more .xlsx E-gradebook files."}, status=400)
    try:
        return JsonResponse(
            process_gradebooks(
                uploads,
                request.POST.get("period", ""),
                request.FILES.get("scoring_workbook"),
                request.POST.get("print_stage", ""),
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
        content, filename = consume_output(job_id, kind)
    except ProcessingError as error:
        return JsonResponse({"error": str(error)}, status=404)
    response = HttpResponse(content)
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    response["Content-Length"] = str(len(content))
    response["Cache-Control"] = "no-store"
    return response
