from threading import local

db = local()


class CountryMiddleware:
    """

    Session db to global loacal for non request views
    """

    def __init__(self, get_response):
        self.get_response = get_response
        # One-time configuration and initialization.

    def __call__(self, request):
        # Code to be executed for each request before
        # the view (and later middleware) are called.
        country = request.session.get('country', 'canada')
        db.country = country
        response = self.get_response(request)

        # Code to be executed for each request/response after
        # the view is called.

        return response
