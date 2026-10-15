export const config = { matcher: ['/((?!_next).*)'] };

export default function middleware(req) {
  const auth = req.headers.get('authorization');

  if (auth) {
    const [scheme, encoded] = auth.split(' ');
    if (scheme === 'Basic' && encoded) {
      const decoded = atob(encoded);
      const [user, pass] = decoded.split(':');
      if (user === 'rto-onboarding' && pass === 'raytanorg') {
        return new Response(null, { status: 200 });
      }
    }
  }

  return new Response('Unauthorised', {
    status: 401,
    headers: {
      'WWW-Authenticate': 'Basic realm="RTO Onboarding Portal", charset="UTF-8"',
    },
  });
}
