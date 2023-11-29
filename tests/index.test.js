const request = require('supertest');
const server = require('../index');

describe('File Upload API', () => {
  let server
  beforeAll(async () => {
    const mod = await import('..');
    server = mod.default;
  });

  afterAll((done) => {
    if (server) {
      server.close(done);
    }
  });

  test('should upload a file', async () => {
    const response = await request(server)
      .post('/upload')
      .attach('file', 'C:/Users/Ruud Naastepad/Downloads/WKB/storyboard/pdf-upload-pc.png'); // Replace with the path to a test file

    expect(response.status).toBe(200);
    expect(response.body.message).toBe('Het bestand is geupload!');
  })

  test('should reset message after 5 seconds', async () => {
    const response = await request(server)
      .post('/upload')
      .attach('file', 'C:/Users/Ruud Naastepad/Downloads/WKB/storyboard/pdf-upload-pc.png'); // Replace with the path to a test file

      // Wait for 6 seconds to ensure the message is reset
    await new Promise(resolve => setTimeout(resolve, 6000));

    // Check if the message is reset to an empty string
    const resetResponse = await request(server).get('/');
    expect(resetResponse.text).toContain('<p id="message"></p>');
  }, 7000);
});
