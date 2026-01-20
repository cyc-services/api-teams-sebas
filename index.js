const express = require('express');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const app = express();
app.use(express.json());

const tenantId = 'a82ffc6f-bd8f-46f6-9b8c-9c793f5d15ce';       
const clientId = 'ad0cf26e-0e92-43a8-aea7-b1e32ba75bae';      
const clientSecret = process.env.CLIENT_SECRET; 
const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
const authProvider = {
    getAccessToken: async () => {
        const token = await credential.getToken('https://graph.microsoft.com/.default');
        return token.token;
    }
};
const graphClient = Client.initWithMiddleware({ authProvider });

app.post('/crear-reunion', async (req, res) => {
    try {
        const { emailOrganizador, asunto } = req.body;

        if (!emailOrganizador) {
            return res.status(400).json({ error: "Falta el emailOrganizador" });
        }

        console.log(`1. Buscando ID del usuario: ${emailOrganizador}...`);
        
        const user = await graphClient.api(`/users/${emailOrganizador}`).get();
        const userId = user.id;

        console.log(`2. ID Encontrado: ${userId}. Creando reunión...`);

        const meetingDetails = {
            startDateTime: new Date().toISOString(),
            endDateTime: new Date(Date.now() + 60 * 60000).toISOString(), 
            subject: asunto || "Reunión de Soporte",
        };

        const meeting = await graphClient
            .api(`/users/${userId}/onlineMeetings`)
            .post(meetingDetails);

        console.log("3. ¡Reunión creada con éxito!");

        res.json({
            exito: true,
            mensaje: "Reunión creada correctamente",
            link_teams: meeting.joinWebUrl
        });

    } catch (error) {
        console.error("❌ ERROR:", error.message);
        res.status(500).json({ 
            error: error.message,
            nota: "Verifica que el correo sea el correcto (@caduse.onmicrosoft.com)"
        });
    }
});

const port = process.env.PORT || 3000;

app.listen(port, () => {
    console.log(`🚀 API corriendo en el puerto ${port}`);
});