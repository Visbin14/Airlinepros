from django.test import TestCase
from account.models import User


class UserModelTest(TestCase):

    def test_string_representation(self):
        user = User(email="basil.jose@fingent.com")
        self.assertTrue(isinstance(user, User))
        self.assertEqual(str(user), user.email)


class LoginTests(TestCase):

    def test_login(self):
        self.user = User.objects.create_user(email='basil.jose@fingent.com', password='pass8442')
        login = self.client.login(email='basil.jose@fingent.com', password='pass8442')
        if login:
            response = self.client.get('/')
            self.assertEqual(response.status_code, 200)
        else:
            response = self.client.get('/login/')
            self.assertEqual(response.status_code, 200)
