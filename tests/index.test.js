const request = require('supertest')
const server = require('../index')
// testbestand om te uploaden:
const filename = 'C:/Users/Ruud Naastepad/Downloads/WKB/storyboard/pdf-upload-pc.png'

describe('File Upload API', () => {

  test('should upload a file', async () => {
    const response = await request(server)
      .post('/upload')
      .attach('file', filename); 

    expect(response.status).toBe(200)
    expect(response.body.message).toBe('Het bestand is geupload!')
  })

  test('should say no files uploaded with status 400', async () => {
    const response = await request(server)
      .post('/upload')
      .attach('file', null); 

    expect(response.status).toBe(400)
    expect(response.body.message).toBe('Er zijn geen bestanden geupload.')
  })

  test('should reset message after 5 seconds', async () => {
    const response = await request(server)
      .post('/upload')
      .attach('file', filename); 

      // Wait for 6 seconds to ensure the message is reset
    await new Promise(resolve => setTimeout(resolve, 6000))

    // Check if the message is reset to an empty string
    const resetResponse = await request(server).get('/')
    expect(resetResponse.text).toContain('<p id="message"></p>')
  }, 7000)

  server.close()
})
