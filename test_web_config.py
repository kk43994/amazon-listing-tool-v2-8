import web.app as web_app


def test_ai_config_test_endpoint_uses_ai_client(monkeypatch):
    client = web_app.app.test_client()

    monkeypatch.setattr(web_app, 'reload_config', lambda: web_app.config)
    monkeypatch.setattr('core.ai_client.ai_text', lambda *args, **kwargs: 'OK')

    response = client.post('/api/config/test', json={})

    assert response.status_code == 200
    payload = response.get_json()
    assert payload['success'] is True
    assert 'AI连接成功' in payload['message']


def test_ai_config_test_endpoint_surfaces_ai_errors(monkeypatch):
    client = web_app.app.test_client()

    monkeypatch.setattr(web_app, 'reload_config', lambda: web_app.config)

    def fake_ai_text(*args, **kwargs):
        raise RuntimeError('boom')

    monkeypatch.setattr('core.ai_client.ai_text', fake_ai_text)

    response = client.post('/api/config/test', json={})

    assert response.status_code == 200
    payload = response.get_json()
    assert payload['success'] is False
    assert 'boom' in payload['message']


def test_reload_config_reads_web_port(monkeypatch):
    monkeypatch.setenv('WEB_PORT', '5003')

    cfg = web_app.reload_config()

    assert cfg.WEB_PORT == 5003
