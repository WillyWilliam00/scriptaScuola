import express from "express";

const router = express.Router();

router.get("/", (_req, res) => {
    res.status(200).json({ message: "OK" });
});

/**
 * Restituisce l'indirizzo IP che il backend attribuisce a chi chiama.
 * Serve a verificare TRUST_PROXY dopo un deploy: deve comparire il proprio IP pubblico,
 * non quello di un proxy (Vercel, Render). Su questo IP si basa il limite ai tentativi di login.
 */
router.get("/ip", (req, res) => {
    res.status(200).json({ ip: req.ip });
});

export default router;