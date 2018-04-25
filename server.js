const liveServer = require('live-server');
const openfinConfigBuilder = require('openfin-config-builder');
const openfinLauncher = require('openfin-launcher');
const path = require('path');

let target;
const configPath = path.resolve('./app.json');
const serverParams = {
    root: path.resolve('.'),
    open: false,
    port: 8080,
    logLevel: 2
};

//Update our config and launch openfin.
function launchOpenFin() {
    openfinConfigBuilder.update({
        startup_app: {
            applicationIcon: target + '/favicon.ico'
        },
        shortcut: {
            icon: target + '/favicon.ico'
        }
    }, configPath)
        .then(openfinLauncher.launchOpenFin({ configPath: configPath }))
        .catch(err => console.log(err));
}


//Start the server server and launch our app.
liveServer.start(serverParams).on('listening', () => {
    const { address, port } = liveServer.server.address();
    target = `http://localhost:${ port }`;
    launchOpenFin();
});
