from django.utils.timezone import now


class SetLastUserActivityMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)

        current_user = request.user

        if current_user.is_authenticated:
            current_user.last_login = now()
            current_user.save()
        return response