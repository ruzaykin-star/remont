export const config = {
  matcher: '/:path*',
};

export default function middleware(request) {
  const expectedUser = process.env.BASIC_AUTH_USER;
  const expectedPassword = process.env.BASIC_AUTH_PASSWORD;

  // If the env vars aren't set (e.g. a preview deploy without them configured),
  // fail open rather than locking everyone out including you.
  if (!expectedUser || !expectedPassword) return;

  const authHeader = request.headers.get('authorization');
  if (authHeader?.startsWith('Basic ')) {
    const [user, password] = atob(authHeader.slice(6)).split(':');
    if (user === expectedUser && password === expectedPassword) return;
  }

  return new Response('Authentication required', {
    status: 401,
    headers: { 'WWW-Authenticate': 'Basic realm="remont-tracker"' },
  });
}
