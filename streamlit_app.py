"""Entrada da hospedagem externa: nunca permite dados temporários em produção."""
import app

app.ONLINE_DEPLOYMENT = True
app.main()
