import { baseUrl } from "./Constants";

export const Fetch = (url) => {
    fetch(baseUrl + url).then(r => {
        if (!r.ok) {
            throw new Error("network request failed");
        }
        return r.json()
    }).then(e => {
        return e;
    }).catch(error => {
        console.error('Error fetching data:', error);
        return 'error';
    });
}