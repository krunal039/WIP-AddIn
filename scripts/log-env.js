console.log("[Pipeline] Current environment variables:");
Object.keys(process.env)
  .sort()
  .forEach((key) => {
    const masked = /(KEY|SECRET|TOKEN|PASSWORD)/i.test(key)
      ? "*****"
      : process.env[key];
    console.log(`  ${key} = ${masked}`);
  });
