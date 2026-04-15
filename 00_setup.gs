function setConfig(){
  const p = PropertiesService.getScriptProperties();
  p.setProperty('DISCORD_WEBHOOK_URL','https://discord.com/api/webhooks/1430141975352442880/2ZHscfY5RCc4tlwqPbaeqV4qwieh2M0J9sqm7g6mSMfp7Og4MAsvPE3LXzAxuQAKdEQ1');
  p.setProperty('SECURITY_TOKEN','tok_bA7m2QxK5eN4pH9tR3vS6yL1zW8dJ0fC2');
  Logger.log('Security props saved');
}
