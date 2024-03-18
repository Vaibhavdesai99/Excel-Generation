module.exports = async (srv) => {
  const db = await cds.connect.to({
    kind: "postgres",
    credentials: {
      host: "localhost",
      port: 5432,
      user: "postgres",
      password: "ramchandra@1999",
      database: "postgres",
      schemas: "public",
    },
  });
  srv.on("READ", "ExcelData", async (req) => {
    try {
      const pipeData = await db.run(SELECT.from("myapp_pipedetails"));
      console.log(pipeData);
      return pipeData;
    } catch (err) {
      console.error("Error fetching users:", err);
      return req.error({ error: "Internal server error" }).code(500);
    }
  });
};
